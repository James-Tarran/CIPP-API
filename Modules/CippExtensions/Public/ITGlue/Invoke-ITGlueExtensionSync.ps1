function Invoke-ITGlueExtensionSync {
    <#
    .FUNCTIONALITY
        Internal
    #>
    param(
        $Configuration,
        $TenantFilter
    )

    try {
        $Conn = Connect-ITGlueAPI -Configuration $Configuration
        $ITGlueConfig = $Configuration.ITGlue

        $Tenant = Get-Tenants -TenantFilter $TenantFilter -IncludeErrors
        $CompanyResult = [PSCustomObject]@{
            Name    = $Tenant.displayName
            Users   = 0
            Devices = 0
            Errors  = [System.Collections.Generic.List[string]]@()
            Logs    = [System.Collections.Generic.List[string]]@()
        }

        # Resolve org mapping and field mappings
        $MappingTable = Get-CIPPTable -TableName 'CippMapping'
        $Mappings = Get-CIPPAzDataTableEntity @MappingTable -Filter "PartitionKey eq 'ITGlueMapping' or PartitionKey eq 'ITGlueFieldMapping'"
        $TenantMap = $Mappings | Where-Object { $_.PartitionKey -eq 'ITGlueMapping' -and $_.RowKey -eq $Tenant.customerId }

        if (!$TenantMap) {
            return 'Tenant not found in ITGlue mapping table'
        }

        $OrgId = $TenantMap.IntegrationId
        $PeopleTypeId = ($Mappings | Where-Object { $_.PartitionKey -eq 'ITGlueFieldMapping' -and $_.RowKey -eq 'Users' }).IntegrationId
        $DeviceTypeId = ($Mappings | Where-Object { $_.PartitionKey -eq 'ITGlueFieldMapping' -and $_.RowKey -eq 'Devices' }).IntegrationId
        $CAPTypeId = ($Mappings | Where-Object { $_.PartitionKey -eq 'ITGlueFieldMapping' -and $_.RowKey -eq 'ConditionalAccessPolicies' }).IntegrationId

        # Get M365 cached data
        $ExtensionCache = Get-CippExtensionReportingData -TenantFilter $Tenant.defaultDomainName -IncludeMailboxes

        # License friendly-name table
        $ModuleBase = Get-Module -Name CippExtensions | Select-Object -ExpandProperty ModuleBase
        $LicTable = Import-Csv (Join-Path $ModuleBase 'ConversionTable.csv')

        # CIPP URL for deep links
        $ConfigTable = Get-CIPPTable -tablename 'Config'
        $CIPPConfigRow = Get-CIPPAzDataTableEntity @ConfigTable -Filter "PartitionKey eq 'InstanceProperties' and RowKey eq 'CIPPURL'"
        $CIPPURL = 'https://{0}' -f $CIPPConfigRow.Value

        $CompanyResult.Logs.Add('Starting ITGlue Extension Sync')

        # Get asset cache table for hash-based change detection
        $ITGlueAssetCache = Get-CIPPTable -tablename 'CacheITGlueAssets'

        # Flatten M365 data
        $Users = $ExtensionCache.Users
        $LicensedUsers = $Users | Where-Object { $null -ne $_.assignedLicenses.skuId } | Sort-Object userPrincipalName
        $Devices = $ExtensionCache.Devices
        $AllRoles = $ExtensionCache.AllRoles
        $AllGroups = $ExtensionCache.Groups
        $Licenses = $ExtensionCache.Licenses
        $Domains = $ExtensionCache.Domains
        $Mailboxes = $ExtensionCache.Mailboxes

        # Get formatted CAPs with human-readable names (uses Invoke-ListConditionalAccessPolicies logic)
        if ($ITGlueConfig.SyncConditionalAccessPolicies -eq $true -and ![string]::IsNullOrEmpty($CAPTypeId)) {
            try {
                $CAPResult = Invoke-ListConditionalAccessPolicies -Request @{ Query = @{ tenantFilter = $TenantFilter } }
                $ConditionalAccessPolicies = $CAPResult.Body.Results
            } catch {
                $CompanyResult.Errors.Add("Failed to fetch formatted CAPs: $_")
                $ConditionalAccessPolicies = @()
            }
        } else {
            $ConditionalAccessPolicies = @()
        }

        $CompanyResult.Users = ($LicensedUsers | Measure-Object).count
        $CompanyResult.Devices = ($Devices | Measure-Object).count

        # Smart auto-create for Conditional Access Policies flexible asset type
        if ($ITGlueConfig.SyncConditionalAccessPolicies -eq $true -and [string]::IsNullOrEmpty($CAPTypeId) -and $ConditionalAccessPolicies -and $ConditionalAccessPolicies.Count -gt 0) {
            try {
                # Search for existing type matching "Conditional Access"
                $AllFlexibleAssetTypes = Invoke-ITGlueRequest -Method GET -Endpoint '/flexible_asset_types' -Headers $Conn.Headers -BaseUrl $Conn.BaseUrl
                $ExistingCAPType = $AllFlexibleAssetTypes | Where-Object { $_.name -like '*Conditional Access*' } | Select-Object -First 1

                if ($ExistingCAPType) {
                    $CAPTypeId = $ExistingCAPType.id
                    $CompanyResult.Logs.Add("Found existing Conditional Access flexible asset type: $($ExistingCAPType.name)")
                } else {
                    # Create flexible asset type with all fields using relationships structure
                    $NewTypeBody = @{
                        data = @{
                            type       = 'flexible-asset-types'
                            attributes = @{
                                name        = 'Conditional Access Policy'
                                description = 'Microsoft 365 Conditional Access Policies synced from CIPP'
                                icon        = 'shield-alt'
                                enabled     = $true
                            }
                            relationships = @{
                                'flexible-asset-fields' = @{
                                    data = @(
                                        @{
                                            type       = 'flexible-asset-fields'
                                            attributes = @{
                                                order          = 1
                                                name           = 'Policy Name'
                                                kind           = 'Text'
                                                required       = $true
                                                'show-in-list' = $true
                                            }
                                        }
                                        @{
                                            type       = 'flexible-asset-fields'
                                            attributes = @{
                                                order          = 2
                                                name           = 'Policy ID'
                                                kind           = 'Text'
                                                required       = $false
                                                'show-in-list' = $false
                                            }
                                        }
                                        @{
                                            type       = 'flexible-asset-fields'
                                            attributes = @{
                                                order          = 3
                                                name           = 'State'
                                                kind           = 'Text'
                                                required       = $false
                                                'show-in-list' = $true
                                            }
                                        }
                                        @{
                                            type       = 'flexible-asset-fields'
                                            attributes = @{
                                                order          = 4
                                                name           = 'Policy Details'
                                                kind           = 'Textbox'
                                                required       = $false
                                                'show-in-list' = $false
                                            }
                                        }
                                        @{
                                            type       = 'flexible-asset-fields'
                                            attributes = @{
                                                order          = 5
                                                name           = 'Raw JSON'
                                                kind           = 'Textbox'
                                                required       = $false
                                                'show-in-list' = $false
                                            }
                                        }
                                    )
                                }
                            }
                        }
                    } | ConvertTo-Json -Depth 20 -Compress

                    $NewType = Invoke-RestMethod -Uri "$($Conn.BaseUrl)/flexible_asset_types" -Method POST -Headers $Conn.Headers -Body $NewTypeBody
                    $CAPTypeId = $NewType.data.id
                    $CompanyResult.Logs.Add("Created new Conditional Access Policy flexible asset type (ID: $CAPTypeId)")
                }

                # Save mapping to database
                $AddMapping = @{
                    PartitionKey    = 'ITGlueFieldMapping'
                    RowKey          = 'ConditionalAccessPolicies'
                    IntegrationId   = "$CAPTypeId"
                    IntegrationName = 'Conditional Access Policy'
                }
                Add-CIPPAzDataTableEntity @MappingTable -Entity $AddMapping -Force
                $CompanyResult.Logs.Add("Saved Conditional Access Policy mapping (ID: $CAPTypeId)")
            } catch {
                $CompanyResult.Errors.Add("Failed to auto-create CAP flexible asset type: $_")
            }
        }

        # Serial exclusion list
        $DefaultSerials = [System.Collections.Generic.List[string]]@(
            'SystemSerialNumber', 'System Serial Number',
            '0123456789', '123456789'
        )
        if ($ITGlueConfig.ExcludeSerials) {
            $DefaultSerials.AddRange(($ITGlueConfig.ExcludeSerials -split ',').Trim())
        }
        $ExcludeSerials = $DefaultSerials

        # Helper: ensure required fields exist in an ITGlue Flexible Asset Type.
        function Add-ITGlueFlexibleAssetFields {
            param(
                $TypeId,
                [array]$FieldsToAdd,  # Array of @{ Name = ''; Kind = 'Textbox'; ShowInList = $false }
                $Conn
            )

            # GET type with its fields included (one call for all fields)
            $TypeResponse = Invoke-RestMethod -Uri "$($Conn.BaseUrl)/flexible_asset_types/$TypeId`?include=flexible_asset_fields" -Method GET -Headers $Conn.Headers
            $IncludedFields = $TypeResponse.included | Where-Object { $_.type -eq 'flexible-asset-fields' }
            $ExistingNames = $IncludedFields | ForEach-Object { $_.attributes.name }

            # Filter to only fields that don't exist
            $NewFields = $FieldsToAdd | Where-Object { $_.Name -notin $ExistingNames }

            if ($NewFields.Count -eq 0) {
                return  # All fields already exist
            }

            # Build complete field list: existing (with IDs) + new fields
            $AllFields = [System.Collections.Generic.List[object]]::new()
            foreach ($F in $IncludedFields) {
                $AllFields.Add([ordered]@{
                    id             = $F.id
                    name           = $F.attributes.name
                    kind           = $F.attributes.kind
                    required       = $F.attributes.required
                    'show-in-list' = $F.attributes.'show-in-list'
                    position       = $F.attributes.position
                })
            }

            foreach ($NewField in $NewFields) {
                $AllFields.Add([ordered]@{
                    name           = $NewField.Name
                    kind           = $NewField.Kind
                    required       = $false
                    'show-in-list' = $NewField.ShowInList
                })
            }

            $PatchBody = @{
                data = @{
                    type       = 'flexible-asset-types'
                    id         = $TypeId
                    attributes = @{
                        'flexible-asset-fields' = @($AllFields)
                    }
                }
            } | ConvertTo-Json -Depth 20 -Compress

            $null = Invoke-RestMethod -Uri "$($Conn.BaseUrl)/flexible_asset_types/$TypeId" -Method PATCH -Headers $Conn.Headers -Body $PatchBody
        }

        # Helper: Convert Out-String output (newline-separated) to comma-separated
        # Used for formatting CAP values from Invoke-ListConditionalAccessPolicies
        function Format-CAPValue($Value) {
            if ([string]::IsNullOrWhiteSpace($Value)) { return '' }
            ($Value.Trim() -split "`n" | Where-Object { $_ -and $_.Trim() } | ForEach-Object { $_.Trim() }) -join ', '
        }

        # USERS — FLEXIBLE ASSETS
        if (![string]::IsNullOrEmpty($PeopleTypeId)) {
            try {
                # Batch field additions into single API call
                Add-ITGlueFlexibleAssetFields -TypeId $PeopleTypeId -FieldsToAdd @(
                    @{ Name = 'Email Address'; Kind = 'Text'; ShowInList = $true }
                    @{ Name = 'Microsoft 365'; Kind = 'Textbox'; ShowInList = $false }
                ) -Conn $Conn

                $ExistingPeopleAssets = Invoke-ITGlueRequest -Method GET -Endpoint '/flexible_assets' -Headers $Conn.Headers -BaseUrl $Conn.BaseUrl -QueryParams @{
                    'filter[flexible_asset_type_id]' = $PeopleTypeId
                    'filter[organization_id]'        = $OrgId
                }

                $UserUpdatedCount = 0
                $UserSkippedCount = 0

                foreach ($User in $LicensedUsers) {
                    try {
                        $UserLicenseNames = foreach ($Lic in $User.assignedLicenses) {
                            $FriendlyName = ($LicTable | Where-Object { $_.SkuId -eq $Lic.skuId }).ProductName
                            if ($FriendlyName) { $FriendlyName } else { $Lic.skuId }
                        }
                        $UserGroups = ($AllGroups | Where-Object { $_.members.id -contains $User.id }).displayName -join ', '
                        $UserRoles  = ($AllRoles  | Where-Object { $_.members.userPrincipalName -contains $User.userPrincipalName }).displayName -join ', '
                        $Mailbox    = $Mailboxes | Where-Object { $_.UPN -eq $User.userPrincipalName } | Select-Object -First 1

                        # Build HTML WITHOUT timestamp first (for hash calculation)
                        $M365HtmlCore = @"
<p><strong>Licenses:</strong> $($UserLicenseNames -join ', ')</p>
<p><strong>Groups:</strong> $(if ($UserGroups) { $UserGroups } else { 'None' })</p>
<p><strong>Admin Roles:</strong> $(if ($UserRoles) { $UserRoles } else { 'None' })</p>
<p><strong>Account Enabled:</strong> $($User.accountEnabled)</p>
<p><strong>Job Title:</strong> $($User.jobTitle)</p>
<p><strong>Department:</strong> $($User.department)</p>
$(if ($Mailbox) { "<p><strong>Mailbox Size:</strong> $($Mailbox.TotalItemSize)</p>" })
<p><a href="$CIPPURL/identity/administration/users?customerId=$($Tenant.customerId)" target="_blank">View in CIPP</a> &nbsp;
<a href="https://entra.microsoft.com/$($Tenant.defaultDomainName)/#view/Microsoft_AAD_UsersAndTenants/UserProfileMenuBlade/~/overview/userId/$($User.id)" target="_blank">View in Entra</a></p>
"@

                        # Hash-based change detection - hash content WITHOUT timestamp
                        $ContentToHash = "$($User.displayName)|$($User.userPrincipalName)|$($User.accountEnabled)|$M365HtmlCore"
                        $NewHash = Get-StringHash -String $ContentToHash

                        # Add timestamp AFTER hashing (for display only)
                        $M365Html = $M365HtmlCore + "`n<p><em>Last updated: $(Get-Date -Format 'yyyy-MM-dd HH:mm') UTC</em></p>"

                        $Traits = @{
                            'name'          = $User.displayName
                            'email-address' = $User.userPrincipalName
                            'microsoft-365' = $M365Html
                        }

                        $ExistingAsset = $ExistingPeopleAssets | Where-Object { $_.traits.'email-address' -eq $User.userPrincipalName } | Select-Object -First 1

                        # Check if content has changed by comparing hashes
                        $NeedsUpdate = $true
                        if ($ExistingAsset) {
                            $CachedAsset = Get-CIPPAzDataTableEntity @ITGlueAssetCache -Filter "PartitionKey eq 'ITGlueUser' and RowKey eq '$($ExistingAsset.id)'"
                            if ($CachedAsset -and $CachedAsset.Hash -eq $NewHash) {
                                $NeedsUpdate = $false
                                $UserSkippedCount++
                            }
                        }

                        if ($NeedsUpdate) {
                            $AssetAttribs = @{
                                'organization-id'        = $OrgId
                                'flexible-asset-type-id' = $PeopleTypeId
                                traits                   = $Traits
                            }

                            if ($ExistingAsset) {
                                $null = Invoke-ITGlueRequest -Method PATCH -Endpoint "/flexible_assets/$($ExistingAsset.id)" -Headers $Conn.Headers -BaseUrl $Conn.BaseUrl -ResourceType 'flexible-assets' -ResourceId $ExistingAsset.id -Attributes $AssetAttribs
                                $AssetId = $ExistingAsset.id
                            } else {
                                $CreatedAsset = Invoke-ITGlueRequest -Method POST -Endpoint '/flexible_assets' -Headers $Conn.Headers -BaseUrl $Conn.BaseUrl -ResourceType 'flexible-assets' -Attributes $AssetAttribs
                                $AssetId = $CreatedAsset[0].id
                            }

                            # Cache the hash to avoid unnecessary updates on next sync
                            $CacheEntry = @{
                                PartitionKey = 'ITGlueUser'
                                RowKey       = [string]$AssetId
                                OrgId        = [string]$OrgId
                                UserUPN      = $User.userPrincipalName
                                Hash         = $NewHash
                            }
                            Add-CIPPAzDataTableEntity @ITGlueAssetCache -Entity $CacheEntry -Force

                            $UserUpdatedCount++
                        }
                    } catch {
                        $CompanyResult.Errors.Add("User FA [$($User.userPrincipalName)]: $_")
                    }
                }

                $CompanyResult.Logs.Add("Users Flexible Assets: $UserUpdatedCount updated, $UserSkippedCount unchanged")
            } catch {
                $CompanyResult.Errors.Add("Users Flexible Assets block failed: $_")
            }
        }

        # USERS — NATIVE CONTACTS
        if ($ITGlueConfig.CreateMissingContacts -eq $true) {
            try {
                $ExistingContacts = Invoke-ITGlueRequest -Method GET -Endpoint '/contacts' -Headers $Conn.Headers -BaseUrl $Conn.BaseUrl -QueryParams @{
                    'filter[organization_id]' = $OrgId
                }

                foreach ($User in $LicensedUsers) {
                    try {
                        # Match by primary email — contacts store emails in a nested array
                        $ExistingContact = $ExistingContacts | Where-Object {
                            ($_.'contact-emails' | Where-Object { $_.value -eq $User.userPrincipalName }) -ne $null
                        } | Select-Object -First 1

                        $ContactAttribs = @{
                            'organization-id' = $OrgId
                            'first-name'      = if ($User.givenName) { $User.givenName } else { $User.displayName }
                            'last-name'       = $User.surname
                            title             = $User.jobTitle
                            'contact-emails'  = @(@{ value = $User.userPrincipalName; primary = $true; 'label-name' = 'Work' })
                        }

                        if ($ExistingContact) {
                            $null = Invoke-ITGlueRequest -Method PATCH -Endpoint "/contacts/$($ExistingContact.id)" -Headers $Conn.Headers -BaseUrl $Conn.BaseUrl -ResourceType 'contacts' -ResourceId $ExistingContact.id -Attributes $ContactAttribs
                        } else {
                            $null = Invoke-ITGlueRequest -Method POST -Endpoint '/contacts' -Headers $Conn.Headers -BaseUrl $Conn.BaseUrl -ResourceType 'contacts' -Attributes $ContactAttribs
                        }
                    } catch {
                        $CompanyResult.Errors.Add("Contact [$($User.userPrincipalName)]: $_")
                    }
                }

                $CompanyResult.Logs.Add("Native Contacts: Processed $($LicensedUsers.Count) users")
            } catch {
                $CompanyResult.Errors.Add("Native Contacts block failed: $_")
            }
        }

        # DEVICES — FLEXIBLE ASSETS
        if (![string]::IsNullOrEmpty($DeviceTypeId)) {
            try {
                Add-ITGlueFlexibleAssetFields -TypeId $DeviceTypeId -FieldsToAdd @(
                    @{ Name = 'Microsoft 365'; Kind = 'Textbox'; ShowInList = $false }
                ) -Conn $Conn

                $ExistingDeviceAssets = Invoke-ITGlueRequest -Method GET -Endpoint '/flexible_assets' -Headers $Conn.Headers -BaseUrl $Conn.BaseUrl -QueryParams @{
                    'filter[flexible_asset_type_id]' = $DeviceTypeId
                    'filter[organization_id]'        = $OrgId
                }

                $SyncDevices = $Devices | Where-Object {
                    $_.serialNumber -notin $ExcludeSerials -and
                    ![string]::IsNullOrWhiteSpace($_.serialNumber) -and
                    $_.managedDeviceOwnerType -eq 'company'
                }

                $DeviceUpdatedCount = 0
                $DeviceSkippedCount = 0

                foreach ($Device in $SyncDevices) {
                    try {
                        # Build HTML WITHOUT timestamp first (for hash calculation)
                        $M365DeviceHtmlCore = @"
<p><strong>Serial:</strong> $($Device.serialNumber)</p>
<p><strong>OS:</strong> $($Device.operatingSystem) $($Device.osVersion)</p>
<p><strong>Manufacturer / Model:</strong> $($Device.manufacturer) $($Device.model)</p>
<p><strong>Compliance:</strong> $($Device.complianceState)</p>
<p><strong>Enrolled:</strong> $($Device.enrolledDateTime)</p>
<p><strong>Last Device Sync:</strong> $($Device.lastSyncDateTime)</p>
<p><strong>Primary User:</strong> $($Device.userDisplayName) ($($Device.userPrincipalName))</p>
<p><a href="$CIPPURL/endpoint/reports/devices?customerId=$($Tenant.customerId)" target="_blank">View in CIPP</a> &nbsp;
<a href="https://intune.microsoft.com/$($Tenant.defaultDomainName)/" target="_blank">Open Intune</a></p>
"@

                        # Hash-based change detection - hash content WITHOUT timestamp
                        $ContentToHash = "$($Device.deviceName)|$($Device.complianceState)|$($Device.lastSyncDateTime)|$M365DeviceHtmlCore"
                        $NewHash = Get-StringHash -String $ContentToHash

                        # Add timestamp AFTER hashing (for display only)
                        $M365DeviceHtml = $M365DeviceHtmlCore + "`n<p><em>Last updated: $(Get-Date -Format 'yyyy-MM-dd HH:mm') UTC</em></p>"

                        $DeviceTraits = @{
                            'name'          = $Device.deviceName
                            'microsoft-365' = $M365DeviceHtml
                        }

                        $ExistingAsset = $ExistingDeviceAssets | Where-Object { $_.traits.name -eq $Device.deviceName } | Select-Object -First 1

                        # Check if content has changed by comparing hashes
                        $NeedsUpdate = $true
                        if ($ExistingAsset) {
                            $CachedAsset = Get-CIPPAzDataTableEntity @ITGlueAssetCache -Filter "PartitionKey eq 'ITGlueDevice' and RowKey eq '$($ExistingAsset.id)'"
                            if ($CachedAsset -and $CachedAsset.Hash -eq $NewHash) {
                                $NeedsUpdate = $false
                                $DeviceSkippedCount++
                            }
                        }

                        if ($NeedsUpdate) {
                            $AssetAttribs = @{
                                'organization-id'        = $OrgId
                                'flexible-asset-type-id' = $DeviceTypeId
                                traits                   = $DeviceTraits
                            }

                            if ($ExistingAsset) {
                                $null = Invoke-ITGlueRequest -Method PATCH -Endpoint "/flexible_assets/$($ExistingAsset.id)" -Headers $Conn.Headers -BaseUrl $Conn.BaseUrl -ResourceType 'flexible-assets' -ResourceId $ExistingAsset.id -Attributes $AssetAttribs
                                $AssetId = $ExistingAsset.id
                            } else {
                                $CreatedAsset = Invoke-ITGlueRequest -Method POST -Endpoint '/flexible_assets' -Headers $Conn.Headers -BaseUrl $Conn.BaseUrl -ResourceType 'flexible-assets' -Attributes $AssetAttribs
                                $AssetId = $CreatedAsset[0].id
                            }

                            # Cache the hash to avoid unnecessary updates on next sync
                            $CacheEntry = @{
                                PartitionKey = 'ITGlueDevice'
                                RowKey       = [string]$AssetId
                                OrgId        = [string]$OrgId
                                DeviceName   = $Device.deviceName
                                Hash         = $NewHash
                            }
                            Add-CIPPAzDataTableEntity @ITGlueAssetCache -Entity $CacheEntry -Force

                            $DeviceUpdatedCount++
                        }
                    } catch {
                        $CompanyResult.Errors.Add("Device FA [$($Device.deviceName)]: $_")
                    }
                }

                $CompanyResult.Logs.Add("Device Flexible Assets: $DeviceUpdatedCount updated, $DeviceSkippedCount unchanged")
            } catch {
                $CompanyResult.Errors.Add("Device Flexible Assets block failed: $_")
            }
        }

        # DEVICES — NATIVE CONFIGURATIONS
        if ($ITGlueConfig.CreateMissingConfigurations -eq $true) {
            try {
                # Cache configuration types for the whole sync run
                $ConfigTypes = Invoke-ITGlueRequest -Method GET -Endpoint '/configuration_types' -Headers $Conn.Headers -BaseUrl $Conn.BaseUrl

                $ExistingConfigs = Invoke-ITGlueRequest -Method GET -Endpoint '/configurations' -Headers $Conn.Headers -BaseUrl $Conn.BaseUrl -QueryParams @{
                    'filter[organization_id]' = $OrgId
                }

                $SyncDevices = $Devices | Where-Object {
                    $_.serialNumber -notin $ExcludeSerials -and
                    ![string]::IsNullOrWhiteSpace($_.serialNumber) -and
                    $_.managedDeviceOwnerType -eq 'company'
                }

                foreach ($Device in $SyncDevices) {
                    try {
                        # Map Intune OS to a common ITGlue configuration type name
                        $ConfigTypeName = switch -Wildcard ($Device.operatingSystem) {
                            'Windows*' { 'Workstation' }
                            'macOS*'   { 'Mac' }
                            'iOS*'     { 'Mobile Device' }
                            'Android*' { 'Mobile Device' }
                            default    { 'Workstation' }
                        }
                        $ConfigType = $ConfigTypes | Where-Object { $_.name -like "*$ConfigTypeName*" } | Select-Object -First 1
                        if (!$ConfigType) { $ConfigType = $ConfigTypes | Select-Object -First 1 }

                        $ConfigAttribs = @{
                            'organization-id'       = $OrgId
                            'configuration-type-id' = $ConfigType.id
                            name                    = $Device.deviceName
                            hostname                = $Device.deviceName
                            'serial-number'         = $Device.serialNumber
                            'operating-system'      = "$($Device.operatingSystem) $($Device.osVersion)"
                            notes                   = "Manufacturer: $($Device.manufacturer)`nModel: $($Device.model)`nCompliance: $($Device.complianceState)`nEnrolled: $($Device.enrolledDateTime)`nLast Sync: $($Device.lastSyncDateTime)`nUser: $($Device.userDisplayName) ($($Device.userPrincipalName))"
                        }

                        # Prefer serial-number match; fall back to device name
                        $ExistingConfig = $ExistingConfigs | Where-Object { $_.'serial-number' -eq $Device.serialNumber } | Select-Object -First 1
                        if (!$ExistingConfig) {
                            $ExistingConfig = $ExistingConfigs | Where-Object { $_.name -eq $Device.deviceName } | Select-Object -First 1
                        }

                        if ($ExistingConfig) {
                            $null = Invoke-ITGlueRequest -Method PATCH -Endpoint "/configurations/$($ExistingConfig.id)" -Headers $Conn.Headers -BaseUrl $Conn.BaseUrl -ResourceType 'configurations' -ResourceId $ExistingConfig.id -Attributes $ConfigAttribs
                        } else {
                            $null = Invoke-ITGlueRequest -Method POST -Endpoint '/configurations' -Headers $Conn.Headers -BaseUrl $Conn.BaseUrl -ResourceType 'configurations' -Attributes $ConfigAttribs
                        }
                    } catch {
                        $CompanyResult.Errors.Add("Config [$($Device.deviceName)]: $_")
                    }
                }

                $CompanyResult.Logs.Add("Native Configurations: Processed $($SyncDevices.Count) devices")
            } catch {
                $CompanyResult.Errors.Add("Native Configurations block failed: $_")
            }
        }

        # CONDITIONAL ACCESS POLICIES — FLEXIBLE ASSETS
        if ($ITGlueConfig.SyncConditionalAccessPolicies -eq $true -and ![string]::IsNullOrEmpty($CAPTypeId) -and $ConditionalAccessPolicies -and $ConditionalAccessPolicies.Count -gt 0) {
            try {
                Add-ITGlueFlexibleAssetFields -TypeId $CAPTypeId -FieldsToAdd @(
                    @{ Name = 'Policy Name'; Kind = 'Text'; ShowInList = $true }
                    @{ Name = 'Policy ID'; Kind = 'Text'; ShowInList = $false }
                    @{ Name = 'State'; Kind = 'Text'; ShowInList = $true }
                    @{ Name = 'Policy Details'; Kind = 'Textbox'; ShowInList = $false }
                    @{ Name = 'Raw JSON'; Kind = 'Textbox'; ShowInList = $false }
                ) -Conn $Conn

                $ExistingCAPAssets = Invoke-ITGlueRequest -Method GET -Endpoint '/flexible_assets' -Headers $Conn.Headers -BaseUrl $Conn.BaseUrl -QueryParams @{
                    'filter[flexible_asset_type_id]' = $CAPTypeId
                    'filter[organization_id]'        = $OrgId
                }

                $UpdatedCount = 0
                $SkippedCount = 0

                foreach ($CAP in $ConditionalAccessPolicies) {
                    try {
                        $StateIcon = switch ($CAP.state) {
                            'enabled' { '✓ Enabled' }
                            'disabled' { '✗ Disabled' }
                            'enabledForReportingButNotEnforced' { '⚠ Report-Only' }
                            default { $CAP.state }
                        }

                        # Build content for hash - ONLY actual policy settings (exclude dates/timestamps)
                        $ContentForHash = @"
State: $StateIcon
Client App Types: $($CAP.clientAppTypes)
Platforms (Include): $($CAP.includePlatforms)
Platforms (Exclude): $($CAP.excludePlatforms)
Locations (Include): $($CAP.includeLocations)
Locations (Exclude): $($CAP.excludeLocations)
Applications (Include): $($CAP.includeApplications)
Applications (Exclude): $($CAP.excludeApplications)
User Actions: $(Format-CAPValue $CAP.includeUserActions)
Auth Context: $(Format-CAPValue $CAP.includeAuthenticationContextClassReferences)
Users (Include): $(Format-CAPValue $CAP.includeUsers)
Users (Exclude): $(Format-CAPValue $CAP.excludeUsers)
Groups (Include): $(Format-CAPValue $CAP.includeGroups)
Groups (Exclude): $(Format-CAPValue $CAP.excludeGroups)
Roles (Include): $(Format-CAPValue $CAP.includeRoles)
Roles (Exclude): $(Format-CAPValue $CAP.excludeRoles)
Operator: $($CAP.grantControlsOperator)
Built-in Controls: $($CAP.builtInControls)
Custom Auth Factors: $($CAP.customAuthenticationFactors)
Terms of Use: $($CAP.termsOfUse)
"@

                        # Hash-based change detection - hash ONLY policy content (not dates or display timestamps)
                        $ContentToHash = "$($CAP.displayName)|$($CAP.state)|$ContentForHash"
                        $NewHash = Get-StringHash -String $ContentToHash

                        # Build full HTML with dates for display (dates NOT in hash)
                        $DetailsHtml = @"
<h4>State: $StateIcon</h4>
<p><strong>Created:</strong> $($CAP.createdDateTime)<br/>
<strong>Modified:</strong> $($CAP.modifiedDateTime)</p>

<h4>Conditions</h4>
<table>
<tr><td><strong>Client App Types</strong></td><td>$($CAP.clientAppTypes)</td></tr>
<tr><td><strong>Platforms (Include)</strong></td><td>$($CAP.includePlatforms)</td></tr>
<tr><td><strong>Platforms (Exclude)</strong></td><td>$($CAP.excludePlatforms)</td></tr>
<tr><td><strong>Locations (Include)</strong></td><td>$($CAP.includeLocations)</td></tr>
<tr><td><strong>Locations (Exclude)</strong></td><td>$($CAP.excludeLocations)</td></tr>
<tr><td><strong>Applications (Include)</strong></td><td>$($CAP.includeApplications)</td></tr>
<tr><td><strong>Applications (Exclude)</strong></td><td>$($CAP.excludeApplications)</td></tr>
<tr><td><strong>User Actions</strong></td><td>$(Format-CAPValue $CAP.includeUserActions)</td></tr>
<tr><td><strong>Auth Context</strong></td><td>$(Format-CAPValue $CAP.includeAuthenticationContextClassReferences)</td></tr>
</table>

<h4>Users & Groups</h4>
<table>
<tr><td><strong>Users (Include)</strong></td><td>$(Format-CAPValue $CAP.includeUsers)</td></tr>
<tr><td><strong>Users (Exclude)</strong></td><td>$(Format-CAPValue $CAP.excludeUsers)</td></tr>
<tr><td><strong>Groups (Include)</strong></td><td>$(Format-CAPValue $CAP.includeGroups)</td></tr>
<tr><td><strong>Groups (Exclude)</strong></td><td>$(Format-CAPValue $CAP.excludeGroups)</td></tr>
<tr><td><strong>Roles (Include)</strong></td><td>$(Format-CAPValue $CAP.includeRoles)</td></tr>
<tr><td><strong>Roles (Exclude)</strong></td><td>$(Format-CAPValue $CAP.excludeRoles)</td></tr>
</table>

<h4>Grant Controls</h4>
<table>
<tr><td><strong>Operator</strong></td><td>$($CAP.grantControlsOperator)</td></tr>
<tr><td><strong>Built-in Controls</strong></td><td>$($CAP.builtInControls)</td></tr>
<tr><td><strong>Custom Auth Factors</strong></td><td>$($CAP.customAuthenticationFactors)</td></tr>
<tr><td><strong>Terms of Use</strong></td><td>$($CAP.termsOfUse)</td></tr>
</table>

<p><em>Last updated: $(Get-Date -Format 'yyyy-MM-dd HH:mm') UTC</em></p>
"@

                        $CAPTraits = @{
                            'policy-name'    = $CAP.displayName
                            'policy-id'      = $CAP.id
                            'state'          = $CAP.state
                            'policy-details' = $DetailsHtml
                            'raw-json'       = $CAP.rawjson
                        }

                        $ExistingAsset = $ExistingCAPAssets | Where-Object { $_.traits.'policy-id' -eq $CAP.id } | Select-Object -First 1

                        # Check if content has changed by comparing hashes
                        $NeedsUpdate = $true
                        if ($ExistingAsset) {
                            $CachedAsset = Get-CIPPAzDataTableEntity @ITGlueAssetCache -Filter "PartitionKey eq 'ITGlueCAP' and RowKey eq '$($ExistingAsset.id)'"
                            if ($CachedAsset -and $CachedAsset.Hash -eq $NewHash) {
                                $NeedsUpdate = $false
                                $SkippedCount++
                            } else {
                                # Debug: Log why hash changed
                                if ($CachedAsset) {
                                    Write-LogMessage -API 'ITGlueSync' -tenant $TenantFilter -message "CAP hash mismatch for $($CAP.displayName): Cached=$($CachedAsset.Hash.Substring(0,8))... New=$($NewHash.Substring(0,8))..." -sev Debug
                                } else {
                                    Write-LogMessage -API 'ITGlueSync' -tenant $TenantFilter -message "CAP no cache found for $($CAP.displayName) (AssetID: $($ExistingAsset.id))" -sev Debug
                                }
                            }
                        }

                        if ($NeedsUpdate) {
                            $AssetAttribs = @{
                                'organization-id'        = $OrgId
                                'flexible-asset-type-id' = $CAPTypeId
                                traits                   = $CAPTraits
                            }

                            if ($ExistingAsset) {
                                $null = Invoke-ITGlueRequest -Method PATCH -Endpoint "/flexible_assets/$($ExistingAsset.id)" -Headers $Conn.Headers -BaseUrl $Conn.BaseUrl -ResourceType 'flexible-assets' -ResourceId $ExistingAsset.id -Attributes $AssetAttribs
                                $AssetId = $ExistingAsset.id
                            } else {
                                $CreatedAsset = Invoke-ITGlueRequest -Method POST -Endpoint '/flexible_assets' -Headers $Conn.Headers -BaseUrl $Conn.BaseUrl -ResourceType 'flexible-assets' -Attributes $AssetAttribs
                                $AssetId = $CreatedAsset[0].id
                            }

                            # Cache the hash to avoid unnecessary updates on next sync
                            $CacheEntry = @{
                                PartitionKey = 'ITGlueCAP'
                                RowKey       = [string]$AssetId
                                OrgId        = [string]$OrgId
                                PolicyId     = $CAP.id
                                Hash         = $NewHash
                            }
                            Add-CIPPAzDataTableEntity @ITGlueAssetCache -Entity $CacheEntry -Force

                            $UpdatedCount++
                        }
                    } catch {
                        $CompanyResult.Errors.Add("CAP FA [$($CAP.displayName)]: $_")
                    }
                }

                # Delete CAP assets that no longer exist in M365
                $CurrentCAPIds = $ConditionalAccessPolicies | ForEach-Object { $_.id }
                $OrphanedAssets = $ExistingCAPAssets | Where-Object { $_.traits.'policy-id' -notin $CurrentCAPIds }
                foreach ($Orphan in $OrphanedAssets) {
                    try {
                        $PolicyName = if ($Orphan.traits.'policy-name') { $Orphan.traits.'policy-name' } else { "ID: $($Orphan.traits.'policy-id')" }
                        $null = Invoke-ITGlueRequest -Method DELETE -Endpoint "/flexible_assets/$($Orphan.id)" -Headers $Conn.Headers -BaseUrl $Conn.BaseUrl
                        $CompanyResult.Logs.Add("Deleted orphaned CAP: $PolicyName")

                        # Remove from cache
                        $CachedAsset = Get-CIPPAzDataTableEntity @ITGlueAssetCache -Filter "PartitionKey eq 'ITGlueCAP' and RowKey eq '$($Orphan.id)'"
                        if ($CachedAsset) {
                            Remove-AzDataTableEntity @ITGlueAssetCache -Entity $CachedAsset -Force
                        }
                    } catch {
                        $PolicyName = if ($Orphan.traits.'policy-name') { $Orphan.traits.'policy-name' } else { "ID: $($Orphan.traits.'policy-id')" }
                        $CompanyResult.Errors.Add("Failed to delete orphaned CAP [$PolicyName]: $_")
                    }
                }

                $CompanyResult.Logs.Add("Conditional Access Policies: $UpdatedCount updated, $SkippedCount unchanged")
            } catch {
                $CompanyResult.Errors.Add("Conditional Access Policies block failed: $_")
            }
        }

        # M365 OVERVIEW — update organisation quick-notes (preserving existing content)
        if ($ITGlueConfig.ImportDomains -eq $true -and $Domains) {
            try {
                $VerifiedDomainList = ($Domains | Where-Object { $_.isVerified -eq $true }).id
                $VerifiedDomains = if ($VerifiedDomainList) {
                    '<ul>' + (($VerifiedDomainList | ForEach-Object { "<li>$_</li>" }) -join '') + '</ul>'
                } else {
                    '<p>None</p>'
                }

                # Build license table rows
                $LicenseRows = if ($Licenses) {
                    foreach ($License in ($Licenses | Where-Object { $_.prepaidUnits.enabled -gt 0 } | Sort-Object -Property skuPartNumber)) {
                        $FriendlyName = ($LicTable | Where-Object { $_.SkuId -eq $License.skuId }).ProductName
                        if (-not $FriendlyName) { $FriendlyName = $License.skuPartNumber }
                        "<tr><td>$FriendlyName</td><td>$($License.consumedUnits) / $($License.prepaidUnits.enabled)</td></tr>"
                    }
                }
                $LicenseTable = if ($LicenseRows) {
                    "<table><thead><tr><th>License</th><th>Used / Total</th></tr></thead><tbody>$($LicenseRows -join '')</tbody></table>"
                } else {
                    '<p>No license data available</p>'
                }

                # CIPP managed section wrapped in a <div> with a class attribute.
                # HTML comments (<!-- -->) are stripped by ITGlue's sanitizer, so we use a real element as our marker instead.
                $CippMarkerStart = '<div class="cipp-managed">'
                $CippMarkerEnd = '</div>'

                $CippSection = @"
$CippMarkerStart
<hr/>
<h3>Microsoft 365 Overview</h3>
<p><strong>Tenant:</strong> $($Tenant.displayName)<br/>
<strong>Tenant ID:</strong> <code>$($Tenant.customerId)</code><br/>
<strong>Default Domain:</strong> $($Tenant.defaultDomainName)</p>

<p><strong>Verified Domains:</strong></p>
$VerifiedDomains

<table>
<tr><td><strong>Licensed Users</strong></td><td>$($CompanyResult.Users)</td></tr>
<tr><td><strong>Managed Devices</strong></td><td>$($CompanyResult.Devices)</td></tr>
</table>

<h4>Licenses</h4>
$LicenseTable

<p><a href="$CIPPURL/tenant/administration/tenants?customerId=$($Tenant.customerId)" target="_blank">View in CIPP</a> |
<a href="https://admin.microsoft.com/Partner/BeginClientSession.aspx?CTID=$($Tenant.customerId)" target="_blank">M365 Admin</a> |
<a href="https://entra.microsoft.com/$($Tenant.defaultDomainName)" target="_blank">Entra Admin</a></p>

<p><em>Last updated: $(Get-Date -Format 'yyyy-MM-dd HH:mm') UTC (CIPP Managed)</em></p>
$CippMarkerEnd
"@

                # Get existing quick-notes from the organization
                $ExistingOrg = Invoke-ITGlueRequest -Method GET -Endpoint "/organizations/$OrgId" -Headers $Conn.Headers -BaseUrl $Conn.BaseUrl -FirstPageOnly
                $ExistingNotes = $ExistingOrg.'quick-notes'

                # ITGlue reformats HTML, so use flexible regex that handles whitespace variations
                if ($ExistingNotes -and $ExistingNotes -match '<div\s+class="cipp-managed">') {
                    # CIPP section exists - replace ALL occurrences (handles duplicates from failed previous syncs)
                    # Use non-capturing group and match any whitespace after opening tag
                    $QuickNotes = $ExistingNotes -replace '(?s)<div\s+class="cipp-managed">.*?</div>\s*', ''
                    # Append fresh CIPP section to cleaned notes
                    if ($QuickNotes.Trim()) {
                        $QuickNotes = $QuickNotes.TrimEnd() + "`n`n" + $CippSection
                    } else {
                        $QuickNotes = $CippSection -replace '<hr/>\s*', ''
                    }
                } elseif ($ExistingNotes -and $ExistingNotes.Trim()) {
                    # No previous CIPP section found - append below existing user content
                    $QuickNotes = $ExistingNotes.TrimEnd() + "`n`n" + $CippSection
                } else {
                    # No existing content, just use CIPP section (without leading hr)
                    $QuickNotes = $CippSection -replace '<hr/>\s*', ''
                }

                $null = Invoke-ITGlueRequest -Method PATCH -Endpoint "/organizations/$OrgId" -Headers $Conn.Headers -BaseUrl $Conn.BaseUrl -ResourceType 'organizations' -ResourceId $OrgId -Attributes @{
                    'quick-notes' = $QuickNotes
                }
                $CompanyResult.Logs.Add("M365 Overview: Updated organisation quick-notes")
            } catch {
                $CompanyResult.Errors.Add("M365 Overview block failed: $_")
            }
        }

        $CompanyResult.Logs.Add('ITGlue Extension Sync complete')
        Write-LogMessage -Message "ITGlue sync complete for $($Tenant.displayName): $($CompanyResult.Users) users, $($CompanyResult.Devices) devices, $($CompanyResult.Errors.Count) errors" -Level Info -tenant $TenantFilter -API 'ITGlueSync'

        return $CompanyResult

    } catch {
        $Message = if ($_.ErrorDetails.Message) {
            Get-NormalizedError -Message $_.ErrorDetails.Message
        } else {
            $_.Exception.message
        }
        Write-LogMessage -Message "ITGlue Extension Sync failed for $TenantFilter : $Message" -Level Error -tenant $TenantFilter -API 'ITGlueSync'
        return "ITGlue sync failed: $Message"
    }
}
