function Invoke-CIPPStandardAppManagementPolicy {
    <#
    .FUNCTIONALITY
        Internal
    .COMPONENT
        (APIName) AppManagementPolicy
    .SYNOPSIS
        (Label) Set Default App Management Policy
    .DESCRIPTION
        (Helptext) Configures the default app management policy to control application and service principal credential restrictions such as password and key credential lifetimes.
        (DocsDescription) Configures the default app management policy to control application and service principal credential restrictions. This includes password addition restrictions, custom password addition, symmetric key addition, and credential lifetime limits for both applications and service principals.
    .NOTES
        CAT
            Entra (AAD) Standards
        TAG
        EXECUTIVETEXT
            Enforces credential restrictions on application registrations and service principals to limit how secrets and certificates are created and how long they remain valid. This reduces the risk of long-lived or unmanaged credentials being used to access your tenant.
        ADDEDCOMPONENT
            {"type":"autoComplete","multiple":false,"creatable":false,"label":"Password Addition","name":"standards.AppManagementPolicy.passwordCredentialsPasswordAddition","options":[{"label":"Enabled","value":"enabled"},{"label":"Disabled","value":"disabled"}]}
            {"type":"autoComplete","multiple":false,"creatable":false,"label":"Custom Password Addition","name":"standards.AppManagementPolicy.passwordCredentialsCustomPasswordAddition","options":[{"label":"Enabled","value":"enabled"},{"label":"Disabled","value":"disabled"}]}
            {"type":"autoComplete","multiple":false,"creatable":false,"label":"Symmetric Key Addition","name":"standards.AppManagementPolicy.keyCredentialsSymmetricKeyAddition","options":[{"label":"Enabled","value":"enabled"},{"label":"Disabled","value":"disabled"}]}
            {"type":"number","label":"Password Credentials Max Lifetime (Days)","name":"standards.AppManagementPolicy.passwordCredentialsMaxLifetime"}
            {"type":"number","label":"Key Credentials Max Lifetime (Days)","name":"standards.AppManagementPolicy.keyCredentialsMaxLifetime"}
        IMPACT
            Medium Impact
        ADDEDDATE
            2026-03-13
        POWERSHELLEQUIVALENT
            Graph API
        RECOMMENDEDBY
        UPDATECOMMENTBLOCK
            Run the Tools\Update-StandardsComments.ps1 script to update this comment block
    .LINK
        https://docs.cipp.app/user-documentation/tenant/standards/list-standards
    #>

    param($Tenant, $Settings)

    # Get current app management policy
    try {
        $CurrentPolicy = New-GraphGetRequest -Uri 'https://graph.microsoft.com/v1.0/policies/defaultAppManagementPolicy' -tenantid $Tenant -AsApp $true
    } catch {
        $ErrorMessage = Get-CippException -Exception $_
        Write-LogMessage -API 'Standards' -tenant $Tenant -message "Failed to get App Management Policy. Error: $($ErrorMessage.NormalizedError)" -sev Error -LogData $ErrorMessage
        return
    }

    # Unwrap autoComplete/number values - frontend may send {label, value} objects or plain values
    $passwordMaxLifetimeDays = $Settings.passwordCredentialsMaxLifetime.value ?? $Settings.passwordCredentialsMaxLifetime
    $keyMaxLifetimeDays = $Settings.keyCredentialsMaxLifetime.value ?? $Settings.keyCredentialsMaxLifetime
    $passwordMaxLifetimeISO = if (-not [string]::IsNullOrWhiteSpace($passwordMaxLifetimeDays) -and $passwordMaxLifetimeDays -ne 'Select a value') { "P${passwordMaxLifetimeDays}D" } else { $null }
    $keyMaxLifetimeISO = if (-not [string]::IsNullOrWhiteSpace($keyMaxLifetimeDays) -and $keyMaxLifetimeDays -ne 'Select a value') { "P${keyMaxLifetimeDays}D" } else { $null }

    # Define desired password credential restrictions from settings
    $PasswordRestrictionDefs = @(
        @{ Setting = 'passwordCredentialsPasswordAddition';       RestrictionType = 'passwordAddition';       UseSettingAsState = $true }
        @{ Setting = 'passwordCredentialsMaxLifetime';            RestrictionType = 'passwordLifetime';       UseSettingAsState = $false; FixedState = 'enabled'; Lifetime = $passwordMaxLifetimeISO }
        @{ Setting = 'passwordCredentialsCustomPasswordAddition'; RestrictionType = 'customPasswordAddition'; UseSettingAsState = $true }
        @{ Setting = 'keyCredentialsSymmetricKeyAddition';        RestrictionType = 'symmetricKeyAddition';   UseSettingAsState = $true }
        @{ Setting = 'keyCredentialsMaxLifetime';                 RestrictionType = 'symmetricKeyLifetime';   UseSettingAsState = $false; FixedState = 'enabled'; Lifetime = $keyMaxLifetimeISO }
    )

    $desiredPasswordCredentials = @(foreach ($def in $PasswordRestrictionDefs) {
        $rawVal = $Settings.($def.Setting)
        $val = $rawVal.value ?? $rawVal
        if (-not [string]::IsNullOrWhiteSpace($val) -and $val -ne 'Select a value') {
            [ordered]@{
                restrictionType                     = $def.RestrictionType
                state                               = if ($def.UseSettingAsState) { $val } else { $def.FixedState }
                maxLifetime                         = if ($def.Lifetime) { $def.Lifetime } else { $null }
                restrictForAppsCreatedAfterDateTime = '0001-01-01T00:00:00Z'
            }
        }
    })

    # Key credentials (asymmetric key lifetime)
    $desiredKeyCredentials = @(
        if ($keyMaxLifetimeISO) {
            [ordered]@{
                restrictionType                     = 'asymmetricKeyLifetime'
                state                               = 'enabled'
                maxLifetime                         = $keyMaxLifetimeISO
                restrictForAppsCreatedAfterDateTime = '0001-01-01T00:00:00Z'
            }
        }
    )

    if ($desiredPasswordCredentials.Count -eq 0 -and $desiredKeyCredentials.Count -eq 0) {
        Write-LogMessage -API 'Standards' -Tenant $Tenant -Message 'AppManagementPolicy: No valid restriction settings were configured.' -Sev Info
        return
    }

    # Build desired state - service principal restrictions mirror application restrictions
    $desiredState = [PSCustomObject]@{
        isEnabled                   = $true
        applicationRestrictions     = [PSCustomObject]@{
            passwordCredentials = $desiredPasswordCredentials
            keyCredentials      = $desiredKeyCredentials
        }
        servicePrincipalRestrictions = [PSCustomObject]@{
            passwordCredentials = $desiredPasswordCredentials
            keyCredentials      = $desiredKeyCredentials
        }
    }

    # Cherry-pick only the properties we manage - New-GraphGetRequest returns a deserialized
    # PSObject that includes extra top-level properties (id, displayName, description, @odata.*)
    $CurrentValue = [PSCustomObject]@{
        isEnabled                   = [bool]$CurrentPolicy.isEnabled
        applicationRestrictions     = $CurrentPolicy.applicationRestrictions
        servicePrincipalRestrictions = $CurrentPolicy.servicePrincipalRestrictions
    }

    $ExpectedValue = [PSCustomObject]@{
        isEnabled                   = $true
        applicationRestrictions     = $desiredState.applicationRestrictions
        servicePrincipalRestrictions = $desiredState.servicePrincipalRestrictions
    }

    # Compare individual properties to avoid JSON key-ordering issues
    $StateIsCorrect = ($CurrentValue.isEnabled -eq $true) -and
        (($CurrentValue.applicationRestrictions.passwordCredentials | ConvertTo-Json -Depth 10 -Compress) -eq ($ExpectedValue.applicationRestrictions.passwordCredentials | ConvertTo-Json -Depth 10 -Compress)) -and
        (($CurrentValue.applicationRestrictions.keyCredentials | ConvertTo-Json -Depth 10 -Compress) -eq ($ExpectedValue.applicationRestrictions.keyCredentials | ConvertTo-Json -Depth 10 -Compress)) -and
        (($CurrentValue.servicePrincipalRestrictions.passwordCredentials | ConvertTo-Json -Depth 10 -Compress) -eq ($ExpectedValue.servicePrincipalRestrictions.passwordCredentials | ConvertTo-Json -Depth 10 -Compress)) -and
        (($CurrentValue.servicePrincipalRestrictions.keyCredentials | ConvertTo-Json -Depth 10 -Compress) -eq ($ExpectedValue.servicePrincipalRestrictions.keyCredentials | ConvertTo-Json -Depth 10 -Compress))

    if ($Settings.remediate -eq $true) {
        if ($StateIsCorrect -eq $true) {
            Write-LogMessage -API 'Standards' -Tenant $Tenant -Message 'App Management Policy is already in the desired state.' -Sev Info
        } else {
            try {
                $GraphRequest = @{
                    tenantID    = $Tenant
                    uri         = 'https://graph.microsoft.com/v1.0/policies/defaultAppManagementPolicy'
                    AsApp       = $true
                    Type        = 'PATCH'
                    ContentType = 'application/json; charset=utf-8'
                    Body        = $desiredState | ConvertTo-Json -Depth 20 -Compress
                }

                $null = New-GraphPostRequest @GraphRequest
                Write-LogMessage -API 'Standards' -Tenant $Tenant -Message 'Updated default app management policy.' -Sev Info
            } catch {
                $ErrorMessage = Get-CippException -Exception $_
                Write-LogMessage -API 'Standards' -Tenant $Tenant -Message "Failed to update default app management policy. Error: $($ErrorMessage.NormalizedError)" -Sev Error -LogData $ErrorMessage
            }
        }
    }

    if ($Settings.alert -eq $true) {
        if ($StateIsCorrect -eq $true) {
            Write-LogMessage -API 'Standards' -Tenant $Tenant -Message 'App Management Policy is configured correctly.' -Sev Info
        } else {
            Write-StandardsAlert -message 'App Management Policy is not configured correctly.' -object $CurrentValue -tenant $Tenant -standardName 'AppManagementPolicy' -standardId $Settings.standardId
            Write-LogMessage -API 'Standards' -Tenant $Tenant -Message 'App Management Policy is not configured correctly.' -Sev Info
        }
    }

    if ($Settings.report -eq $true) {
        Set-CIPPStandardsCompareField -FieldName 'standards.AppManagementPolicy' -CurrentValue $CurrentValue -ExpectedValue $ExpectedValue -TenantFilter $Tenant
        Add-CIPPBPAField -FieldName 'AppManagementPolicy' -FieldValue $StateIsCorrect -StoreAs bool -Tenant $Tenant
    }
}
