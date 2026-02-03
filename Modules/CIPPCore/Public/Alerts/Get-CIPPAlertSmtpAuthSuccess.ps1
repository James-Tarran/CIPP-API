function Get-CIPPAlertSmtpAuthSuccess {
    [CmdletBinding()]
    param (
        # REQUIRED: absorbs the ScheduledCommand object
        [Parameter(ValueFromPipeline = $true)]
        [object]$ScheduledCommand,

        [Parameter(Mandatory = $true)]
        $TenantFilter
    )

    begin {
        $Results = @()
    }

    process {
        # Intentionally empty
        # This stops the orchestrator object from flowing further
    }

    end {
        $uri = "https://graph.microsoft.com/v1.0/auditLogs/signIns?`$filter=clientAppUsed eq 'Authenticated SMTP' and status/errorCode eq 0"
        $SignIns = New-GraphGetRequest -Uri $uri -TenantId $TenantFilter

        if ($SignIns.value) {
            $Results = @(
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

            Write-AlertTrace `
                -cmdletName $MyInvocation.MyCommand `
                -tenantFilter $TenantFilter `
                -data $Results
        }

        return $Results
    }
}
