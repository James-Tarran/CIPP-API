function Sync-ITGlueConditionalAccessPolicies {
    <#
    .FUNCTIONALITY
        Internal
    .SYNOPSIS
        Syncs Conditional Access Policies to ITGlue Flexible Assets.
    #>
    param(
        $CAPTypeId,
        $OrgId,
        $Conn,
        $ConditionalAccessPolicies,
        $ITGlueAssetCache,
        $TenantFilter,
        $CIPPURL,
        $Tenant
    )

    $Result = @{
        UpdatedCount = 0
        SkippedCount = 0
        Errors = [System.Collections.Generic.List[string]]@()
        Logs = [System.Collections.Generic.List[string]]@()
    }

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
User Actions: $(Format-ITGlueCAPValue $CAP.includeUserActions)
Auth Context: $(Format-ITGlueCAPValue $CAP.includeAuthenticationContextClassReferences)
Users (Include): $(Format-ITGlueCAPValue $CAP.includeUsers)
Users (Exclude): $(Format-ITGlueCAPValue $CAP.excludeUsers)
Groups (Include): $(Format-ITGlueCAPValue $CAP.includeGroups)
Groups (Exclude): $(Format-ITGlueCAPValue $CAP.excludeGroups)
Roles (Include): $(Format-ITGlueCAPValue $CAP.includeRoles)
Roles (Exclude): $(Format-ITGlueCAPValue $CAP.excludeRoles)
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
<tr><td><strong>User Actions</strong></td><td>$(Format-ITGlueCAPValue $CAP.includeUserActions)</td></tr>
<tr><td><strong>Auth Context</strong></td><td>$(Format-ITGlueCAPValue $CAP.includeAuthenticationContextClassReferences)</td></tr>
</table>

<h4>Users & Groups</h4>
<table>
<tr><td><strong>Users (Include)</strong></td><td>$(Format-ITGlueCAPValue $CAP.includeUsers)</td></tr>
<tr><td><strong>Users (Exclude)</strong></td><td>$(Format-ITGlueCAPValue $CAP.excludeUsers)</td></tr>
<tr><td><strong>Groups (Include)</strong></td><td>$(Format-ITGlueCAPValue $CAP.includeGroups)</td></tr>
<tr><td><strong>Groups (Exclude)</strong></td><td>$(Format-ITGlueCAPValue $CAP.excludeGroups)</td></tr>
<tr><td><strong>Roles (Include)</strong></td><td>$(Format-ITGlueCAPValue $CAP.includeRoles)</td></tr>
<tr><td><strong>Roles (Exclude)</strong></td><td>$(Format-ITGlueCAPValue $CAP.excludeRoles)</td></tr>
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
                        $Result.SkippedCount++
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

                    $Result.UpdatedCount++
                }
            } catch {
                $Result.Errors.Add("CAP FA [$($CAP.displayName)]: $_")
            }
        }

        # Delete CAP assets that no longer exist in M365
        $CurrentCAPIds = $ConditionalAccessPolicies | ForEach-Object { $_.id }
        $OrphanedAssets = $ExistingCAPAssets | Where-Object { $_.traits.'policy-id' -notin $CurrentCAPIds }
        foreach ($Orphan in $OrphanedAssets) {
            try {
                $PolicyName = if ($Orphan.traits.'policy-name') { $Orphan.traits.'policy-name' } else { "ID: $($Orphan.traits.'policy-id')" }
                $null = Invoke-ITGlueRequest -Method DELETE -Endpoint "/flexible_assets/$($Orphan.id)" -Headers $Conn.Headers -BaseUrl $Conn.BaseUrl
                $Result.Logs.Add("Deleted orphaned CAP: $PolicyName")

                # Remove from cache
                $CachedAsset = Get-CIPPAzDataTableEntity @ITGlueAssetCache -Filter "PartitionKey eq 'ITGlueCAP' and RowKey eq '$($Orphan.id)'"
                if ($CachedAsset) {
                    Remove-AzDataTableEntity @ITGlueAssetCache -Entity $CachedAsset -Force
                }
            } catch {
                $PolicyName = if ($Orphan.traits.'policy-name') { $Orphan.traits.'policy-name' } else { "ID: $($Orphan.traits.'policy-id')" }
                $Result.Errors.Add("Failed to delete orphaned CAP [$PolicyName]: $_")
            }
        }

        $Result.Logs.Add("Conditional Access Policies: $($Result.UpdatedCount) updated, $($Result.SkippedCount) unchanged")
    } catch {
        $Result.Errors.Add("Conditional Access Policies block failed: $_")
    }

    return $Result
}
