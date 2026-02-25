function Get-ITGlueFieldMapping {
    <#
    .FUNCTIONALITY
        Internal
    #>
    [CmdletBinding()]
    param (
        $CIPPMapping
    )

    $Mappings = Get-ExtensionMapping -Extension 'ITGlueField'

    $CIPPFieldHeaders = @(
        [PSCustomObject]@{
            Title       = 'ITGlue Flexible Asset Types'
            FieldType   = 'FlexibleAssetTypes'
            Description = 'Map your ITGlue Flexible Asset Types to the CIPP data type. A "Microsoft 365" rich-text field will be added to the layout if it does not already exist. Native Contacts (users) and Configurations (devices) are controlled separately by the toggle settings.'
        }
    )

    $CIPPFields = @(
        [PSCustomObject]@{
            FieldName  = 'Users'
            FieldLabel = 'Flexible Asset Type for M365 Users'
            FieldType  = 'FlexibleAssetTypes'
        }
        [PSCustomObject]@{
            FieldName  = 'Devices'
            FieldLabel = 'Flexible Asset Type for M365 Devices'
            FieldType  = 'FlexibleAssetTypes'
        }
    )

    $Table = Get-CIPPTable -TableName Extensionsconfig
    try {
        $Configuration = (Get-CIPPAzDataTableEntity @Table).config | ConvertFrom-Json -ea stop
        $Conn = Connect-ITGlueAPI -Configuration $Configuration

        try {
            $RawTypes = Invoke-ITGlueRequest -Method GET -Endpoint '/flexible_asset_types' -Headers $Conn.Headers -BaseUrl $Conn.BaseUrl
            $FlexibleAssetTypes = $RawTypes | Select-Object @{Name = 'FieldType'; Expression = { 'FlexibleAssetTypes' } },
                @{Name = 'value'; Expression = { "$($_.id)" } },
                @{Name = 'name'; Expression = { $_.name } }
        } catch {
            $Message = $_.Exception.Message
            Write-Warning "Could not get ITGlue Flexible Asset Types, error: $Message"
            Write-LogMessage -Message "Could not get ITGlue Flexible Asset Types, error: $Message" -Level Error -tenant 'CIPP' -API 'ITGlueMapping'
            $FlexibleAssetTypes = @([PSCustomObject]@{ FieldType = 'FlexibleAssetTypes'; name = "Could not get Flexible Asset Types: $Message"; value = '-1' })
        }
    } catch {
        $Message = $_.Exception.Message
        Write-Warning "Could not connect to ITGlue, error: $Message"
        Write-LogMessage -Message "Could not connect to ITGlue, error: $Message" -Level Error -tenant 'CIPP' -API 'ITGlueMapping'
        $FlexibleAssetTypes = @([PSCustomObject]@{ FieldType = 'FlexibleAssetTypes'; name = "Could not connect to ITGlue: $Message"; value = '-1' })
    }

    $Unset = [PSCustomObject]@{
        name  = '--- Do not synchronize ---'
        value = $null
        type  = 'unset'
    }

    return [PSCustomObject]@{
        CIPPFields        = $CIPPFields
        CIPPFieldHeaders  = $CIPPFieldHeaders
        IntegrationFields = @($Unset) + @($FlexibleAssetTypes)
        Mappings          = @($Mappings)
    }
}
