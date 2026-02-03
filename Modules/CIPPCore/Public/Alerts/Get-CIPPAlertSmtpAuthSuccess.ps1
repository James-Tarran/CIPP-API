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
            return @()
        }

        $AlertData = @(
            $SignIns.value | ForEach-Object {
                [pscustomobject]@{
                    Tenant            = $TenantFilter
                    UserPrincipalName = $_.userPrincipalName
                    CreatedDateTime   = $_.createdDateTime
                    ClientAppUsed     = $_.clientAppUsed
                    IPAddress         = $_.ipAddress
                    ErrorCode         = $_.status.errorCode
                }
            }
        )

        # Log for dedupe / history
        Write-AlertTrace `
            -cmdletName $MyInvocation.MyCommand `
            -tenantFilter $TenantFilter `
            -data $AlertData

        # 🔑 THIS IS THE CRITICAL LINE
        return $AlertData
    }
    catch {
        return @()
    }
}
