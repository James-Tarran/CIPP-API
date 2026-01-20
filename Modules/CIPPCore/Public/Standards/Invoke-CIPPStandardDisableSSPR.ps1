function Invoke-CIPPStandardDisableSSPR {
    <#
    .FUNCTIONALITY
        Internal
    .COMPONENT
        (APIName) DisableSSPR
    .SYNOPSIS
        (Label) Disable SSPR
    .DESCRIPTION
        (Helptext) Disables the self service password reset feature for all accounts in a tenant.
        (DocsDescription) Administrators should follow more stringent password reset procedures rather than self-service options.
    .NOTES
        CAT
            Global Standards
        TAG
        EXECUTIVETEXT
            Administrators should not be allowed to use self-service password reset (SSPR) for enhanced security. Admin accounts require more stringent security controls and should follow formal password reset procedures that involve additional verification steps rather than self-service options. This ensures that administrative account password resets are properly audited and controlled.
        ADDEDCOMPONENT
        IMPACT
            Low Impact
        ADDEDDATE
            2024-06-05
        POWERSHELLEQUIVALENT
            Update-MgPolicyAuthorizationPolicy -AllowedToUseSspr:$false
        RECOMMENDEDBY
        UPDATECOMMENTBLOCK
            Run the Tools\Update-StandardsComments.ps1 script to update this comment block
    .LINK
        https://docs.cipp.app/user-documentation/tenant/standards/list-standards
    #>

    param ($Tenant, $Settings)
    ##$Rerun -Type Standard -Tenant $Tenant -Settings $Settings 'DisableSSPR'

    $Uri = 'https://graph.microsoft.com/v1.0/policies/authorizationPolicy?$select=allowedToUseSSPR'
    try {
        $CurrentState = New-GraphGetRequest -Uri $Uri -tenantid $Tenant
    } catch {
        $ErrorMessage = Get-CippException -Exception $_
        Write-LogMessage -API 'Standards' -tenant $Tenant -message "Could not get CurrentState for Pronouns. Error: $($ErrorMessage.NormalizedError)" -sev Error -LogData $ErrorMessage
        Return
    }

    if ($Settings.remediate -eq $true) {
        Write-Host 'Time to remediate'

        if ($CurrentState.allowedToUseSSPR -eq $false) {
            Write-LogMessage -API 'Standards' -tenant $tenant -message 'Pronouns are already enabled.' -sev Info
        } else {
            $CurrentState.allowedToUseSSPR = $false
            try {
                $Body = ConvertTo-Json -InputObject $CurrentState -Depth 10 -Compress
                $null = New-GraphPostRequest -Uri $Uri -tenantid $Tenant -Body $Body -type PATCH -AsApp $true
                Write-LogMessage -API 'Standards' -tenant $tenant -message 'Disabled SSPR for everyone.' -sev Info
            } catch {
                $ErrorMessage = Get-CippException -Exception $_
                Write-LogMessage -API 'Standards' -tenant $tenant -message "Failed to disable SSPR. Error: $($ErrorMessage.NormalizedError)" -sev Error -LogData $ErrorMessage
            }
        }
    }

    if ($Settings.alert -eq $true) {

        if ($CurrentState.allowedToUseSSPR -eq $false) {
            Write-LogMessage -API 'Standards' -tenant $tenant -message 'SSPR is disabled for all users.' -sev Info
        } else {
            Write-StandardsAlert -message 'SSPR is not disabled for all users' -object $CurrentState -tenant $tenant -standardName 'DisableSSPR' -standardId $Settings.standardId
            Write-LogMessage -API 'Standards' -tenant $tenant -message 'SSPR is not disabled for all users.' -sev Info
        }
    }

    if ($Settings.report -eq $true) {
        Set-CIPPStandardsCompareField -FieldName 'standards.DisableSSPR' -FieldValue $CurrentState.allowedToUseSSPR -Tenant $tenant
        Add-CIPPBPAField -FieldName 'SSPRDisabled' -FieldValue $CurrentState.allowedToUseSSPR -StoreAs bool -Tenant $tenant
    }
}
