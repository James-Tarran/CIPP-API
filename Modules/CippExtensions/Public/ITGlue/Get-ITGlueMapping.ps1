function Get-ITGlueMapping {
    <#
    .FUNCTIONALITY
        Internal
    #>
    [CmdletBinding()]
    param (
        $CIPPMapping
    )

    $ExtensionMappings = Get-ExtensionMapping -Extension 'ITGlue'
    $Tenants = Get-Tenants -IncludeErrors

    $Mappings = foreach ($Mapping in $ExtensionMappings) {
        $Tenant = $Tenants | Where-Object { $_.RowKey -eq $Mapping.RowKey }
        if ($Tenant) {
            [PSCustomObject]@{
                TenantId        = $Tenant.customerId
                Tenant          = $Tenant.displayName
                TenantDomain    = $Tenant.defaultDomainName
                IntegrationId   = $Mapping.IntegrationId
                IntegrationName = $Mapping.IntegrationName
            }
        }
    }

    $Table = Get-CIPPTable -TableName Extensionsconfig
    try {
        $Configuration = (Get-CIPPAzDataTableEntity @Table).config | ConvertFrom-Json -ea stop
        $Conn = Connect-ITGlueAPI -Configuration $Configuration
        $ITGlueOrgs = Invoke-ITGlueRequest -Method GET -Endpoint '/organizations' -Headers $Conn.Headers -BaseUrl $Conn.BaseUrl
    } catch {
        $Message = if ($_.ErrorDetails.Message) {
            Get-NormalizedError -Message $_.ErrorDetails.Message
        } else {
            $_.Exception.message
        }
        Write-LogMessage -Message "Could not get ITGlue Organizations, error: $Message" -Level Error -tenant 'CIPP' -API 'ITGlueMapping'
        $ITGlueOrgs = @([PSCustomObject]@{ name = "Could not get ITGlue Organizations, error: $Message"; id = '-1' })
    }

    $Companies = $ITGlueOrgs | ForEach-Object {
        [PSCustomObject]@{
            name  = $_.name
            value = "$($_.id)"
        }
    }

    return [PSCustomObject]@{
        Companies = @($Companies)
        Mappings  = @($Mappings)
    }
}
