function Set-ITGlueMapping {
    <#
    .FUNCTIONALITY
        Internal
    .SYNOPSIS
        Replaces all ITGlue tenant-to-organisation mappings in the CippMapping table.
        $Request.Body is an array sent by the frontend, each item containing:
          - TenantId        : the CIPP tenant's customerId (used as RowKey)
          - IntegrationId   : the ITGlue organisation ID (value from the org dropdown)
          - IntegrationName : the ITGlue organisation name (display name from the org dropdown)
    #>
    [CmdletBinding()]
    param (
        $CIPPMapping,
        $APIName,
        $Request
    )

    # Remove all existing mappings for this extension
    Get-CIPPAzDataTableEntity @CIPPMapping -Filter "PartitionKey eq 'ITGlueMapping'" | ForEach-Object {
        Remove-AzDataTableEntity -Force @CIPPMapping -Entity $_
    }

    foreach ($Mapping in $Request.Body) {
        $AddObject = @{
            PartitionKey    = 'ITGlueMapping'
            RowKey          = "$($Mapping.TenantId)"
            IntegrationId   = "$($Mapping.IntegrationId)"
            IntegrationName = "$($Mapping.IntegrationName)"
        }
        Add-CIPPAzDataTableEntity @CIPPMapping -Entity $AddObject -Force
        Write-LogMessage -API $APIName -headers $Request.Headers -message "Added ITGlue mapping for $($Mapping.IntegrationName)." -Sev 'Info'
    }

    return [PSCustomObject]@{ Results = 'Successfully edited ITGlue mapping table.' }
}
