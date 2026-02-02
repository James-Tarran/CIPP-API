function Get-CIPPAlertSmtpAuthSuccess {
    <#
    .FUNCTIONALITY
        Entrypoint – Check sign-in logs for SMTP AUTH with success status
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)]
        [Alias('input')]
        $InputValue,

        $TenantFilter
    )

    try {
        # Graph API endpoint for sign-ins
        $uri = "https://graph.microsoft.com/v1.0/auditLogs/signIns?`$filter=clientAppUsed eq 'Authenticated SMTP' and status/errorCode eq 0"

        # Call Graph API for the given tenant
        $SignIns = New-GraphGetRequest -Uri $uri -TenantId $TenantFilter

        if (-not $SignIns.value) {
            return
        }

        # FORCE array output (important for CIPP AddRange)
        $AlertData = @(
            $SignIns.value | Select-Object `
                userPrincipalName,
                createdDateTime,
                clientAppUsed,
                ipAddress,
                status,
                @{ Name = 'Tenant'; Expression = { $TenantFilter } }
        )

        # Write results into the alert pipeline
        Write-AlertTrace `
            -CmdletName $MyInvocation.MyCommand `
            -TenantFilter $TenantFilter `
            -Data $AlertData

    }
    catch {
        # Optional logging if desired
        # Write-AlertMessage -Tenant $TenantFilter -Message "Failed SMTP AUTH sign-in query: $($_.Exception.Message)"
    }
}
