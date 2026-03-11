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
        $CATypeId = ($Mappings | Where-Object { $_.PartitionKey -eq 'ITGlueFieldMapping' -and $_.RowKey -eq 'ConditionalAccessPolicies' }).IntegrationId

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

        # Flatten M365 data
        $Users = $ExtensionCache.Users
        $LicensedUsers = $Users | Where-Object { $null -ne $_.assignedLicenses.skuId } | Sort-Object userPrincipalName
        $Devices = $ExtensionCache.Devices
        $AllRoles = $ExtensionCache.AllRoles
        $AllGroups = $ExtensionCache.Groups
        $Licenses = $ExtensionCache.Licenses
        $Domains = $ExtensionCache.Domains
        $Mailboxes = $ExtensionCache.Mailboxes

        $CompanyResult.Users = ($LicensedUsers | Measure-Object).count
        $CompanyResult.Devices = ($Devices | Measure-Object).count

        # Serial exclusion list
        $DefaultSerials = [System.Collections.Generic.List[string]]@(
            'SystemSerialNumber', 'To Be Filled By O.E.M.', 'System Serial Number',
            '0123456789', '123456789', 'TobefilledbyO.E.M.'
        )
        if ($ITGlueConfig.ExcludeSerials) {
            $DefaultSerials.AddRange(($ITGlueConfig.ExcludeSerials -split ',').Trim())
        }
        $ExcludeSerials = $DefaultSerials

        # ─────────────────────────────────────────────────────────────────────
        # Helper: ensure required fields exist in an ITGlue Flexible Asset Type.
        # Uses raw Invoke-RestMethod to handle the JSON:API 'included' response.
        # ─────────────────────────────────────────────────────────────────────
        function Add-ITGlueFlexibleAssetField {
            param($TypeId, $FieldName, $FieldKind = 'Textbox', $ShowInList = $false, $Conn)

            # GET type with its fields included
            $TypeResponse = Invoke-RestMethod -Uri "$($Conn.BaseUrl)/flexible_asset_types/$TypeId`?include=flexible_asset_fields" -Method GET -Headers $Conn.Headers
            $IncludedFields = $TypeResponse.included | Where-Object { $_.type -eq 'flexible-asset-fields' }
            $ExistingNames = $IncludedFields | ForEach-Object { $_.attributes.name }

            if ($ExistingNames -contains $FieldName) {
                return  # Already exists, nothing to do
            }

            # Build complete field list: existing (with IDs) + new field
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
            $AllFields.Add([ordered]@{
                name           = $FieldName
                kind           = $FieldKind
                required       = $false
                'show-in-list' = $ShowInList
            })

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

        # ─────────────────────────────────────────────────────────────────────
        # USERS — FLEXIBLE ASSETS
        # ─────────────────────────────────────────────────────────────────────
        if (![string]::IsNullOrEmpty($PeopleTypeId)) {
            try {
                Add-ITGlueFlexibleAssetField -TypeId $PeopleTypeId -FieldName 'Email Address' -FieldKind 'Text' -ShowInList $true -Conn $Conn
                Add-ITGlueFlexibleAssetField -TypeId $PeopleTypeId -FieldName 'Microsoft 365'  -FieldKind 'Textbox' -ShowInList $false -Conn $Conn

                $ExistingPeopleAssets = Invoke-ITGlueRequest -Method GET -Endpoint '/flexible_assets' -Headers $Conn.Headers -BaseUrl $Conn.BaseUrl -QueryParams @{
                    'filter[flexible_asset_type_id]' = $PeopleTypeId
                    'filter[organization_id]'        = $OrgId
                }

                foreach ($User in $LicensedUsers) {
                    try {
                        $UserLicenseNames = foreach ($Lic in $User.assignedLicenses) {
                            $FriendlyName = ($LicTable | Where-Object { $_.SkuId -eq $Lic.skuId }).ProductName
                            if ($FriendlyName) { $FriendlyName } else { $Lic.skuId }
                        }
                        $UserGroups = ($AllGroups | Where-Object { $_.members.id -contains $User.id }).displayName -join ', '
                        $UserRoles  = ($AllRoles  | Where-Object { $_.members.userPrincipalName -contains $User.userPrincipalName }).displayName -join ', '
                        $Mailbox    = $Mailboxes | Where-Object { $_.UPN -eq $User.userPrincipalName } | Select-Object -First 1

                        $M365Html = @"
<p><strong>Licenses:</strong> $($UserLicenseNames -join ', ')</p>
<p><strong>Groups:</strong> $(if ($UserGroups) { $UserGroups } else { 'None' })</p>
<p><strong>Admin Roles:</strong> $(if ($UserRoles) { $UserRoles } else { 'None' })</p>
<p><strong>Account Enabled:</strong> $($User.accountEnabled)</p>
<p><strong>Job Title:</strong> $($User.jobTitle)</p>
<p><strong>Department:</strong> $($User.department)</p>
$(if ($Mailbox) { "<p><strong>Mailbox Size:</strong> $($Mailbox.TotalItemSize)</p>" })
<p><a href="$CIPPURL/identity/administration/users?customerId=$($Tenant.customerId)" target="_blank">View in CIPP</a> &nbsp;
<a href="https://entra.microsoft.com/$($Tenant.defaultDomainName)/#view/Microsoft_AAD_UsersAndTenants/UserProfileMenuBlade/~/overview/userId/$($User.id)" target="_blank">View in Entra</a></p>
<p><em>Last updated: $(Get-Date -Format 'yyyy-MM-dd HH:mm') UTC</em></p>
"@

                        $Traits = @{
                            'name'          = $User.displayName
                            'email-address' = $User.userPrincipalName
                            'microsoft-365' = $M365Html
                        }

                        $ExistingAsset = $ExistingPeopleAssets | Where-Object { $_.'email-address' -eq $User.userPrincipalName } | Select-Object -First 1

                        $AssetAttribs = @{
                            'organization-id'        = $OrgId
                            'flexible-asset-type-id' = $PeopleTypeId
                            traits                   = $Traits
                        }

                        if ($ExistingAsset) {
                            $null = Invoke-ITGlueRequest -Method PATCH -Endpoint "/flexible_assets/$($ExistingAsset.id)" -Headers $Conn.Headers -BaseUrl $Conn.BaseUrl -ResourceType 'flexible-assets' -ResourceId $ExistingAsset.id -Attributes $AssetAttribs
                        } else {
                            $null = Invoke-ITGlueRequest -Method POST -Endpoint '/flexible_assets' -Headers $Conn.Headers -BaseUrl $Conn.BaseUrl -ResourceType 'flexible-assets' -Attributes $AssetAttribs
                        }
                        Start-Sleep -Milliseconds 100
                    } catch {
                        $CompanyResult.Errors.Add("User FA [$($User.userPrincipalName)]: $_")
                    }
                }

                $CompanyResult.Logs.Add("Users Flexible Assets: Processed $($LicensedUsers.Count) users")
            } catch {
                $CompanyResult.Errors.Add("Users Flexible Assets block failed: $_")
            }
        }

        # ─────────────────────────────────────────────────────────────────────
        # USERS — NATIVE CONTACTS
        # ─────────────────────────────────────────────────────────────────────
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
                        Start-Sleep -Milliseconds 100
                    } catch {
                        $CompanyResult.Errors.Add("Contact [$($User.userPrincipalName)]: $_")
                    }
                }

                $CompanyResult.Logs.Add("Native Contacts: Processed $($LicensedUsers.Count) users")
            } catch {
                $CompanyResult.Errors.Add("Native Contacts block failed: $_")
            }
        }

        # ─────────────────────────────────────────────────────────────────────
        # DEVICES — FLEXIBLE ASSETS
        # ─────────────────────────────────────────────────────────────────────
        if (![string]::IsNullOrEmpty($DeviceTypeId)) {
            try {
                Add-ITGlueFlexibleAssetField -TypeId $DeviceTypeId -FieldName 'Microsoft 365' -FieldKind 'Textbox' -ShowInList $false -Conn $Conn

                $ExistingDeviceAssets = Invoke-ITGlueRequest -Method GET -Endpoint '/flexible_assets' -Headers $Conn.Headers -BaseUrl $Conn.BaseUrl -QueryParams @{
                    'filter[flexible_asset_type_id]' = $DeviceTypeId
                    'filter[organization_id]'        = $OrgId
                }

                $SyncDevices = $Devices | Where-Object {
                    $_.serialNumber -notin $ExcludeSerials -and
                    ![string]::IsNullOrWhiteSpace($_.serialNumber) -and
                    $_.managedDeviceOwnerType -eq 'company'
                }

                foreach ($Device in $SyncDevices) {
                    try {
                        $M365DeviceHtml = @"
<p><strong>Serial:</strong> $($Device.serialNumber)</p>
<p><strong>OS:</strong> $($Device.operatingSystem) $($Device.osVersion)</p>
<p><strong>Manufacturer / Model:</strong> $($Device.manufacturer) $($Device.model)</p>
<p><strong>Compliance:</strong> $($Device.complianceState)</p>
<p><strong>Enrolled:</strong> $($Device.enrolledDateTime)</p>
<p><strong>Last Device Sync:</strong> $($Device.lastSyncDateTime)</p>
<p><strong>Primary User:</strong> $($Device.userDisplayName) ($($Device.userPrincipalName))</p>
<p><a href="$CIPPURL/endpoint/reports/devices?customerId=$($Tenant.customerId)" target="_blank">View in CIPP</a> &nbsp;
<a href="https://intune.microsoft.com/$($Tenant.defaultDomainName)/" target="_blank">Open Intune</a></p>
<p><em>Last updated: $(Get-Date -Format 'yyyy-MM-dd HH:mm') UTC</em></p>
"@

                        $DeviceTraits = @{
                            'name'          = $Device.deviceName
                            'microsoft-365' = $M365DeviceHtml
                        }

                        $ExistingAsset = $ExistingDeviceAssets | Where-Object { $_.name -eq $Device.deviceName } | Select-Object -First 1

                        $AssetAttribs = @{
                            'organization-id'        = $OrgId
                            'flexible-asset-type-id' = $DeviceTypeId
                            traits                   = $DeviceTraits
                        }

                        if ($ExistingAsset) {
                            $null = Invoke-ITGlueRequest -Method PATCH -Endpoint "/flexible_assets/$($ExistingAsset.id)" -Headers $Conn.Headers -BaseUrl $Conn.BaseUrl -ResourceType 'flexible-assets' -ResourceId $ExistingAsset.id -Attributes $AssetAttribs
                        } else {
                            $null = Invoke-ITGlueRequest -Method POST -Endpoint '/flexible_assets' -Headers $Conn.Headers -BaseUrl $Conn.BaseUrl -ResourceType 'flexible-assets' -Attributes $AssetAttribs
                        }
                        Start-Sleep -Milliseconds 100
                    } catch {
                        $CompanyResult.Errors.Add("Device FA [$($Device.deviceName)]: $_")
                    }
                }

                $CompanyResult.Logs.Add("Device Flexible Assets: Processed $($SyncDevices.Count) devices")
            } catch {
                $CompanyResult.Errors.Add("Device Flexible Assets block failed: $_")
            }
        }

        # ─────────────────────────────────────────────────────────────────────
        # DEVICES — NATIVE CONFIGURATIONS
        # ─────────────────────────────────────────────────────────────────────
        if ($ITGlueConfig.CreateMissingConfigurations -eq $true) {
            try {
                # Cache configuration types for the whole sync run (one API call)
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
                        Start-Sleep -Milliseconds 100
                    } catch {
                        $CompanyResult.Errors.Add("Config [$($Device.deviceName)]: $_")
                    }
                }

                $CompanyResult.Logs.Add("Native Configurations: Processed $($SyncDevices.Count) devices")
            } catch {
                $CompanyResult.Errors.Add("Native Configurations block failed: $_")
            }
        }

        # ─────────────────────────────────────────────────────────────────────
        # CONDITIONAL ACCESS POLICIES — FLEXIBLE ASSETS
        # ─────────────────────────────────────────────────────────────────────
        if (![string]::IsNullOrEmpty($CATypeId)) {
            try {
                $CAPolicies = $ExtensionCache.ConditionalAccess
                if ($CAPolicies -and $CAPolicies.Count -gt 0) {
                    # Fetch reference data for ID resolution
                    $Requests = @(
                        @{ id = 'namedLocations'; url = 'identity/conditionalAccess/namedLocations'; method = 'GET' }
                        @{ id = 'roleDefinitions'; url = 'roleManagement/directory/roleDefinitions?$select=id,displayName'; method = 'GET' }
                        @{ id = 'servicePrincipals'; url = 'servicePrincipals?$top=999&$select=appId,displayName'; method = 'GET' }
                        @{ id = 'applications'; url = 'applications?$top=999&$select=appId,displayName'; method = 'GET' }
                    )
                    $BulkResults = New-GraphBulkRequest -Requests $Requests -tenantid $Tenant.defaultDomainName -asapp $true

                    $NamedLocations = ($BulkResults | Where-Object { $_.id -eq 'namedLocations' }).body.value
                    $RoleDefinitions = ($BulkResults | Where-Object { $_.id -eq 'roleDefinitions' }).body.value
                    $ServicePrincipals = ($BulkResults | Where-Object { $_.id -eq 'servicePrincipals' }).body.value
                    $Applications = ($BulkResults | Where-Object { $_.id -eq 'applications' }).body.value

                    # Helper functions for ID resolution
                    function Get-ResolvedLocationName {
                        param($Id)
                        if ($Id -eq 'All') { return 'All' }
                        if ($Id -eq 'AllTrusted') { return 'All trusted locations' }
                        $Location = $NamedLocations | Where-Object { $_.id -eq $Id } | Select-Object -First 1
                        if ($Location) { return $Location.displayName }
                        return $Id
                    }

                    function Get-ResolvedRoleName {
                        param($Id)
                        if ($Id -eq 'All') { return 'All' }
                        $Role = $RoleDefinitions | Where-Object { $_.id -eq $Id } | Select-Object -First 1
                        if ($Role) { return $Role.displayName }
                        return $Id
                    }

                    function Get-ResolvedAppName {
                        param($Id)
                        if ($Id -eq 'All') { return 'All' }
                        if ($Id -eq 'None') { return 'None' }
                        if ($Id -eq 'Office365') { return 'Office 365' }
                        if ($Id -eq 'MicrosoftAdminPortals') { return 'Microsoft Admin Portals' }
                        $App = $ServicePrincipals | Where-Object { $_.appId -eq $Id } | Select-Object -First 1
                        if (!$App) { $App = $Applications | Where-Object { $_.appId -eq $Id } | Select-Object -First 1 }
                        if (!$App) { $App = $Applications | Where-Object { $_.id -eq $Id } | Select-Object -First 1 }
                        if ($App) { return $App.displayName }
                        return $Id
                    }

                    function Get-ResolvedUserOrGroupName {
                        param($Id)
                        if ($Id -eq 'All') { return 'All users' }
                        if ($Id -eq 'None') { return 'None' }
                        if ($Id -eq 'GuestsOrExternalUsers') { return 'All guest and external users' }
                        $User = $Users | Where-Object { $_.id -eq $Id } | Select-Object -First 1
                        if ($User) { return "$($User.displayName) ($($User.userPrincipalName))" }
                        $Group = $AllGroups | Where-Object { $_.id -eq $Id } | Select-Object -First 1
                        if ($Group) { return $Group.displayName }
                        return $Id
                    }

                    # Get existing CA Policy assets for this org
                    $ExistingCAAssets = Invoke-ITGlueRequest -Method GET -Endpoint '/flexible_assets' -Headers $Conn.Headers -BaseUrl $Conn.BaseUrl -QueryParams @{
                        'filter[flexible_asset_type_id]' = $CATypeId
                        'filter[organization_id]'        = $OrgId
                    }

                    foreach ($Policy in $CAPolicies) {
                        try {
                            # Determine Users include type
                            $UsersIncludeType = 'None'
                            if ($Policy.conditions.users.includeUsers -contains 'All') {
                                $UsersIncludeType = 'All users'
                            } elseif ($Policy.conditions.users.includeUsers -contains 'GuestsOrExternalUsers' -or $Policy.conditions.users.includeGuestsOrExternalUsers) {
                                $UsersIncludeType = 'All guest and external users'
                            } elseif ($Policy.conditions.users.includeUsers -or $Policy.conditions.users.includeGroups -or $Policy.conditions.users.includeRoles) {
                                $UsersIncludeType = 'Select users and groups'
                            }

                            # Determine Users exclude type
                            $UsersExcludeType = 'None'
                            if ($Policy.conditions.users.excludeUsers -contains 'GuestsOrExternalUsers' -or $Policy.conditions.users.excludeGuestsOrExternalUsers) {
                                $UsersExcludeType = 'All guest and external users'
                            } elseif ($Policy.conditions.users.excludeUsers -or $Policy.conditions.users.excludeGroups -or $Policy.conditions.users.excludeRoles) {
                                $UsersExcludeType = 'Select users and groups'
                            }

                            # Resolve included users and groups
                            $IncludedUsersAndGroups = @()
                            if ($Policy.conditions.users.includeUsers) {
                                $IncludedUsersAndGroups += $Policy.conditions.users.includeUsers | Where-Object { $_ -notin @('All', 'GuestsOrExternalUsers', 'None') } | ForEach-Object { Get-ResolvedUserOrGroupName -Id $_ }
                            }
                            if ($Policy.conditions.users.includeGroups) {
                                $IncludedUsersAndGroups += $Policy.conditions.users.includeGroups | ForEach-Object { Get-ResolvedUserOrGroupName -Id $_ }
                            }

                            # Resolve excluded users and groups
                            $ExcludedUsersAndGroups = @()
                            if ($Policy.conditions.users.excludeUsers) {
                                $ExcludedUsersAndGroups += $Policy.conditions.users.excludeUsers | Where-Object { $_ -notin @('All', 'GuestsOrExternalUsers', 'None') } | ForEach-Object { Get-ResolvedUserOrGroupName -Id $_ }
                            }
                            if ($Policy.conditions.users.excludeGroups) {
                                $ExcludedUsersAndGroups += $Policy.conditions.users.excludeGroups | ForEach-Object { Get-ResolvedUserOrGroupName -Id $_ }
                            }

                            # Resolve directory roles
                            $IncludedRoles = ($Policy.conditions.users.includeRoles | ForEach-Object { Get-ResolvedRoleName -Id $_ }) -join "`n"
                            $ExcludedRoles = ($Policy.conditions.users.excludeRoles | ForEach-Object { Get-ResolvedRoleName -Id $_ }) -join "`n"

                            # Determine Target resources include type
                            $TargetResourcesInclude = 'None'
                            if ($Policy.conditions.applications.includeApplications -contains 'All') {
                                $TargetResourcesInclude = 'All cloud apps'
                            } elseif ($Policy.conditions.applications.includeUserActions) {
                                $TargetResourcesInclude = 'User actions'
                            } elseif ($Policy.conditions.applications.includeAuthenticationContextClassReferences) {
                                $TargetResourcesInclude = 'Authentication context'
                            } elseif ($Policy.conditions.applications.includeApplications) {
                                $TargetResourcesInclude = 'Select apps'
                            }

                            # Resolve applications
                            $SelectApps = ($Policy.conditions.applications.includeApplications | Where-Object { $_ -notin @('All', 'None') } | ForEach-Object { Get-ResolvedAppName -Id $_ }) -join "`n"
                            $ExcludedApps = ($Policy.conditions.applications.excludeApplications | ForEach-Object { Get-ResolvedAppName -Id $_ }) -join "`n"

                            # Determine Network include type
                            $NetworkInclude = 'Any network or location'
                            if ($Policy.conditions.locations.includeLocations -contains 'All') {
                                $NetworkInclude = 'Any network or location'
                            } elseif ($Policy.conditions.locations.includeLocations -contains 'AllTrusted') {
                                $NetworkInclude = 'All trusted networks and locations'
                            } elseif ($Policy.conditions.locations.includeLocations) {
                                $NetworkInclude = 'Selected networks and locations'
                            }

                            # Determine Network exclude type
                            $NetworkExclude = 'None'
                            if ($Policy.conditions.locations.excludeLocations -contains 'AllTrusted') {
                                $NetworkExclude = 'All trusted networks and locations'
                            } elseif ($Policy.conditions.locations.excludeLocations) {
                                $NetworkExclude = 'Selected networks and locations'
                            }

                            # Resolve locations
                            $IncludeLocations = ($Policy.conditions.locations.includeLocations | Where-Object { $_ -notin @('All', 'AllTrusted') } | ForEach-Object { Get-ResolvedLocationName -Id $_ }) -join "`n"
                            $ExcludeLocations = ($Policy.conditions.locations.excludeLocations | Where-Object { $_ -notin @('All', 'AllTrusted') } | ForEach-Object { Get-ResolvedLocationName -Id $_ }) -join "`n"

                            # Determine Grant or Block
                            $GrantOrBlock = if ($Policy.grantControls.builtInControls -contains 'block') { 'Block access' } else { 'Grant access' }

                            # Authentication strength
                            $AuthStrength = $Policy.grantControls.authenticationStrength.displayName

                            # Device filter
                            $DeviceFilterMode = 'Not configured'
                            if ($Policy.conditions.devices.deviceFilter.mode) {
                                $DeviceFilterMode = switch ($Policy.conditions.devices.deviceFilter.mode) {
                                    'include' { 'Include' }
                                    'exclude' { 'Exclude' }
                                    default { 'Not configured' }
                                }
                            }

                            # Session controls
                            $SignInFreqEnabled = [bool]$Policy.sessionControls.signInFrequency.isEnabled
                            $SignInFreqValue = $Policy.sessionControls.signInFrequency.value
                            $SignInFreqType = $Policy.sessionControls.signInFrequency.type
                            if ($Policy.sessionControls.signInFrequency.frequencyInterval -eq 'everyTime') {
                                $SignInFreqType = 'everyTime'
                            }

                            $PersistentBrowserEnabled = [bool]$Policy.sessionControls.persistentBrowser.isEnabled
                            $PersistentBrowserMode = $Policy.sessionControls.persistentBrowser.mode

                            $CAEMode = 'Not configured'
                            if ($Policy.sessionControls.continuousAccessEvaluation.mode) {
                                $CAEMode = switch ($Policy.sessionControls.continuousAccessEvaluation.mode) {
                                    'disabled' { 'Disabled' }
                                    'strictEnforcement' { 'Strictly enforced' }
                                    default { 'Not configured' }
                                }
                            }

                            # Build traits hashtable
                            $Traits = @{
                                'policy-name'                                     = $Policy.displayName
                                'policy-description'                              = $Policy.description
                                'policy-state'                                    = $Policy.state

                                # Users
                                'users-include'                                   = $UsersIncludeType
                                'included-guest-or-external-roles'                = if ($Policy.conditions.users.includeGuestsOrExternalUsers) { ($Policy.conditions.users.includeGuestsOrExternalUsers | ConvertTo-Json -Depth 5) } else { '' }
                                'included-directory-roles'                        = $IncludedRoles
                                'included-users-and-groups'                       = ($IncludedUsersAndGroups -join "`n")
                                'users-exclude'                                   = $UsersExcludeType
                                'excluded-guest-or-external-roles'                = if ($Policy.conditions.users.excludeGuestsOrExternalUsers) { ($Policy.conditions.users.excludeGuestsOrExternalUsers | ConvertTo-Json -Depth 5) } else { '' }
                                'excluded-directory-roles'                        = $ExcludedRoles
                                'excluded-users-and-groups'                       = ($ExcludedUsersAndGroups -join "`n")

                                # Target resources
                                'target-resources-include'                        = $TargetResourcesInclude
                                'select-apps'                                     = $SelectApps
                                'user-actions'                                    = ($Policy.conditions.applications.includeUserActions -join "`n")
                                'authentication-context'                          = ($Policy.conditions.applications.includeAuthenticationContextClassReferences -join "`n")
                                'target-resources-excluded-cloud-apps'            = $ExcludedApps

                                # Network
                                'network-include'                                 = $NetworkInclude
                                'network-include-selected-networks-and-locations' = $IncludeLocations
                                'network-exclude'                                 = $NetworkExclude
                                'network-exclude-selected-networks-and-locations' = $ExcludeLocations

                                # User risk
                                'user-risk-high'                                  = $Policy.conditions.userRiskLevels -contains 'high'
                                'user-risk-medium'                                = $Policy.conditions.userRiskLevels -contains 'medium'
                                'user-risk-low'                                   = $Policy.conditions.userRiskLevels -contains 'low'

                                # Sign-in risk
                                'sign-in-risk-high'                               = $Policy.conditions.signInRiskLevels -contains 'high'
                                'sign-in-risk-medium'                             = $Policy.conditions.signInRiskLevels -contains 'medium'
                                'sign-in-risk-low'                                = $Policy.conditions.signInRiskLevels -contains 'low'
                                'sign-in-risk-no-risk'                            = $Policy.conditions.signInRiskLevels -contains 'none'

                                # Insider risk
                                'insider-risk-elevated'                           = $Policy.conditions.insiderRiskLevels -contains 'elevated'
                                'insider-risk-moderate'                           = $Policy.conditions.insiderRiskLevels -contains 'moderate'
                                'insider-risk-minor'                              = $Policy.conditions.insiderRiskLevels -contains 'minor'

                                # Device platforms - Include
                                'include-android'                                 = $Policy.conditions.platforms.includePlatforms -contains 'android'
                                'include-ios'                                     = $Policy.conditions.platforms.includePlatforms -contains 'iOS'
                                'include-windows'                                 = $Policy.conditions.platforms.includePlatforms -contains 'windows'
                                'include-macos'                                   = $Policy.conditions.platforms.includePlatforms -contains 'macOS'
                                'include-linux'                                   = $Policy.conditions.platforms.includePlatforms -contains 'linux'
                                'include-windows-phone'                           = $Policy.conditions.platforms.includePlatforms -contains 'windowsPhone'

                                # Device platforms - Exclude
                                'exclude-android'                                 = $Policy.conditions.platforms.excludePlatforms -contains 'android'
                                'exclude-ios'                                     = $Policy.conditions.platforms.excludePlatforms -contains 'iOS'
                                'exclude-windows'                                 = $Policy.conditions.platforms.excludePlatforms -contains 'windows'
                                'exclude-macos'                                   = $Policy.conditions.platforms.excludePlatforms -contains 'macOS'
                                'exclude-linux'                                   = $Policy.conditions.platforms.excludePlatforms -contains 'linux'
                                'exclude-windows-phone'                           = $Policy.conditions.platforms.excludePlatforms -contains 'windowsPhone'

                                # Client apps
                                'client-apps-configured'                          = [bool]$Policy.conditions.clientAppTypes
                                'browser'                                         = $Policy.conditions.clientAppTypes -contains 'browser'
                                'mobile-apps-and-desktop-clients'                 = $Policy.conditions.clientAppTypes -contains 'mobileAppsAndDesktopClients'
                                'exchange-activesync-clients'                     = $Policy.conditions.clientAppTypes -contains 'exchangeActiveSync'
                                'other-clients'                                   = $Policy.conditions.clientAppTypes -contains 'other'

                                # Device filter
                                'device-filter-mode'                              = $DeviceFilterMode
                                'device-filter-rule'                              = $Policy.conditions.devices.deviceFilter.rule

                                # Grant controls
                                'grant-or-block'                                  = $GrantOrBlock
                                'grant-controls-operator'                         = $Policy.grantControls.operator
                                'require-multifactor-authentication'              = $Policy.grantControls.builtInControls -contains 'mfa'
                                'require-authentication-strength'                 = $AuthStrength
                                'require-device-to-be-marked-as-compliant'        = $Policy.grantControls.builtInControls -contains 'compliantDevice'
                                'require-microsoft-entra-hybrid-joined-device'    = $Policy.grantControls.builtInControls -contains 'domainJoinedDevice'
                                'require-approved-client-app'                     = $Policy.grantControls.builtInControls -contains 'approvedApplication'
                                'require-app-protection-policy'                   = $Policy.grantControls.builtInControls -contains 'compliantApplication'
                                'require-password-change'                         = $Policy.grantControls.builtInControls -contains 'passwordChange'
                                'terms-of-use'                                    = ($Policy.grantControls.termsOfUse -join "`n")

                                # Session controls
                                'use-app-enforced-restrictions'                   = [bool]$Policy.sessionControls.applicationEnforcedRestrictions.isEnabled
                                'use-conditional-access-app-control'              = [bool]$Policy.sessionControls.cloudAppSecurity.isEnabled
                                'conditional-access-app-control-type'             = $Policy.sessionControls.cloudAppSecurity.cloudAppSecurityType
                                'sign-in-frequency-enabled'                       = $SignInFreqEnabled
                                'sign-in-frequency-value'                         = "$SignInFreqValue"
                                'sign-in-frequency-type'                          = $SignInFreqType
                                'persistent-browser-session-enabled'              = $PersistentBrowserEnabled
                                'persistent-browser-session-mode'                 = $PersistentBrowserMode
                                'continuous-access-evaluation'                    = $CAEMode
                                'disable-resilience-defaults'                     = [bool]$Policy.sessionControls.disableResilienceDefaults
                                'secure-sign-in-session'                          = [bool]$Policy.sessionControls.secureSignInSession.isEnabled

                                # Metadata
                                'policy-id'                                       = $Policy.id
                                'created-date'                                    = if ($Policy.createdDateTime) { $Policy.createdDateTime } else { '' }
                                'modified-date'                                   = if ($Policy.modifiedDateTime) { $Policy.modifiedDateTime } else { '' }
                                'cipp-link'                                       = "$CIPPURL/tenant/conditional/list-policies?customerId=$($Tenant.customerId)"
                                'entra-link'                                      = "https://entra.microsoft.com/$($Tenant.defaultDomainName)/#view/Microsoft_AAD_ConditionalAccess/PolicyBlade/policyId/$($Policy.id)"
                                'last-synced'                                     = (Get-Date -Format 'yyyy-MM-dd HH:mm') + ' UTC'
                            }

                            # Find existing asset by policy ID
                            $ExistingAsset = $ExistingCAAssets | Where-Object { $_.'policy-id' -eq $Policy.id } | Select-Object -First 1

                            $AssetAttribs = @{
                                'organization-id'        = $OrgId
                                'flexible-asset-type-id' = $CATypeId
                                traits                   = $Traits
                            }

                            if ($ExistingAsset) {
                                $null = Invoke-ITGlueRequest -Method PATCH -Endpoint "/flexible_assets/$($ExistingAsset.id)" -Headers $Conn.Headers -BaseUrl $Conn.BaseUrl -ResourceType 'flexible-assets' -ResourceId $ExistingAsset.id -Attributes $AssetAttribs
                            } else {
                                $null = Invoke-ITGlueRequest -Method POST -Endpoint '/flexible_assets' -Headers $Conn.Headers -BaseUrl $Conn.BaseUrl -ResourceType 'flexible-assets' -Attributes $AssetAttribs
                            }
                            Start-Sleep -Milliseconds 100
                        } catch {
                            $CompanyResult.Errors.Add("CA Policy FA [$($Policy.displayName)]: $_")
                        }
                    }

                    $CompanyResult.Logs.Add("CA Policy Flexible Assets: Processed $($CAPolicies.Count) policies")
                } else {
                    $CompanyResult.Logs.Add('CA Policy Flexible Assets: No policies found in cache')
                }
            } catch {
                $CompanyResult.Errors.Add("CA Policy Flexible Assets block failed: $_")
            }
        }

        # ─────────────────────────────────────────────────────────────────────
        # M365 OVERVIEW — update organisation quick-notes (preserving existing content)
        # ─────────────────────────────────────────────────────────────────────
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
                # HTML comments (<!-- -->) are stripped by ITGlue's sanitizer, so we
                # use a real element as our marker instead.
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

                if ($ExistingNotes -and $ExistingNotes -match '\(CIPP Managed\)') {
                    # Content-based match: use GREEDY .* to capture everything from
                    # the first M365 Overview heading to the LAST (CIPP Managed) stamp,
                    # removing all duplicate CIPP sections in one sweep.
                    $QuickNotes = $ExistingNotes -replace '(?s)(<hr\s*/?>)?\s*<h3>Microsoft 365 Overview</h3>.*(CIPP Managed)</em></p>', $CippSection
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
