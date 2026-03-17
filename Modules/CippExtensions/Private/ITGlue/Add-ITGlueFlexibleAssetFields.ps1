function Add-ITGlueFlexibleAssetFields {
    <#
    .FUNCTIONALITY
        Internal
    .SYNOPSIS
        Ensures required fields exist in an ITGlue Flexible Asset Type.
    #>
    param(
        $TypeId,
        [array]$FieldsToAdd,  # Array of @{ Name = ''; Kind = 'Textbox'; ShowInList = $false }
        $Conn
    )

    # GET type with its fields included (one call for all fields)
    $TypeResponse = Invoke-RestMethod -Uri "$($Conn.BaseUrl)/flexible_asset_types/$TypeId`?include=flexible_asset_fields" -Method GET -Headers $Conn.Headers
    $IncludedFields = $TypeResponse.included | Where-Object { $_.type -eq 'flexible-asset-fields' }
    $ExistingNames = $IncludedFields | ForEach-Object { $_.attributes.name }

    # Filter to only fields that don't exist
    $NewFields = $FieldsToAdd | Where-Object { $_.Name -notin $ExistingNames }

    if ($NewFields.Count -eq 0) {
        return  # All fields already exist
    }

    # Build complete field list: existing (with IDs) + new fields
    $AllFields = [System.Collections.Generic.List[object]]::new()
    foreach ($F in $IncludedFields) {
        $AllFields.Add([ordered]@{
            id             = $F.id
            name           = $F.attributes.name
            kind           = $F.attributes.kind
            required       = $F.attributes.required
            'show-in-list' = $F.attributes.'show-in-list'
            position       = $F.attributes.position
        })
    }

    foreach ($NewField in $NewFields) {
        $AllFields.Add([ordered]@{
            name           = $NewField.Name
            kind           = $NewField.Kind
            required       = $false
            'show-in-list' = $NewField.ShowInList
        })
    }

    $PatchBody = @{
        data = @{
            type       = 'flexible-asset-types'
            id         = $TypeId
            attributes = @{
                'flexible-asset-fields' = @($AllFields)
            }
        }
    } | ConvertTo-Json -Depth 20 -Compress

    $null = Invoke-RestMethod -Uri "$($Conn.BaseUrl)/flexible_asset_types/$TypeId" -Method PATCH -Headers $Conn.Headers -Body $PatchBody
}
