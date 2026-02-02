function Get-CIPPAlertSmtpAuthSuccess {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        $TenantFilter
    )

    try {
        $uri = "https://graph.microsoft.com/v1.0/auditLogs/signIns?`$filter=clientAppUsed eq 'Authenticated SMTP' and status/errorCode eq 0"

        $SignIns = New-GraphGetRequest -Uri $uri -TenantId $TenantFilter

        if (-not $SignIns.value) {
            return
        }

        $AlertData = @(
            $SignIns.value | Select-Object `
                userPrincipalName,
                createdDateTime,
                clientAppUsed,
                ipAddress,
                status,
                @{ Name = 'Tenant'; Expression = { $TenantFilter } }
        )

        Write-AlertTrace `
            -cmdletName $MyInvocation.MyCommand `
            -tenantFilter $TenantFilter `
            -data $AlertData
    }
    catch {
        return
    }
}
