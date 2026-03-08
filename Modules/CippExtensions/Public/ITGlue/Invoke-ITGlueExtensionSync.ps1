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

                if ($ExistingNotes) {
                    # Try matching our <div class="cipp-managed"> wrapper first
                    if ($ExistingNotes -match '<div class="cipp-managed">') {
                        $QuickNotes = $ExistingNotes -replace '(?s)<div class="cipp-managed">.*?</div>', $CippSection
                    }
                    # Legacy: match old HTML comment markers (in case ITGlue kept them)
                    elseif ($ExistingNotes -match '<!-- CIPP-MANAGED-START -->') {
                        $QuickNotes = $ExistingNotes -replace '(?s)<!-- CIPP-MANAGED-START -->.*?<!-- CIPP-MANAGED-END -->', $CippSection
                    }
                    # Fallback: detect orphaned CIPP content by matching on the heading
                    # and the "(CIPP Managed)" timestamp that we always write
                    elseif ($ExistingNotes -match '<h3>Microsoft 365 Overview</h3>' -and $ExistingNotes -match '\(CIPP Managed\)') {
                        $QuickNotes = $ExistingNotes -replace '(?s)(<hr\s*/?>)?\s*<h3>Microsoft 365 Overview</h3>.*?\(CIPP Managed\)</em></p>', $CippSection
                    }
                    # No previous CIPP section found - append below existing user content
                    elseif ($ExistingNotes.Trim()) {
                        $QuickNotes = $ExistingNotes.TrimEnd() + "`n`n" + $CippSection
                    }
                    else {
                        $QuickNotes = $CippSection -replace '<hr/>\s*', ''
                    }
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
