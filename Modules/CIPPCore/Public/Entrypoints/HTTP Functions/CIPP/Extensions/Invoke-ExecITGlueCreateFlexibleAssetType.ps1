Function Invoke-ExecITGlueCreateFlexibleAssetType {
  <#
    .FUNCTIONALITY
        Entrypoint
    .ROLE
        CIPP.Extension.ReadWrite
    #>
  [CmdletBinding()]
  param($Request, $TriggerMetadata)

  $APIName = $Request.Params.CIPPEndpoint
  $Headers = $Request.Headers

  try {
    $AssetType = $Request.Body.AssetType

    if ([string]::IsNullOrEmpty($AssetType)) {
      throw "AssetType parameter is required"
    }

    # Only support ConditionalAccessPolicies for now
    if ($AssetType -ne 'ConditionalAccessPolicies') {
      throw "Unsupported asset type: $AssetType"
    }

    # Get ITGlue configuration
    $Table = Get-CIPPTable -TableName Extensionsconfig
    $Configuration = (Get-CIPPAzDataTableEntity @Table).config | ConvertFrom-Json -ErrorAction Stop
    $Conn = Connect-ITGlueAPI -Configuration $Configuration

    if (-not $Conn) {
      throw "Failed to connect to ITGlue API. Please check your configuration."
    }

    # Search for existing type matching "Conditional Access"
    $AllFlexibleAssetTypes = Invoke-ITGlueRequest -Method GET -Endpoint '/flexible_asset_types' -Headers $Conn.Headers -BaseUrl $Conn.BaseUrl
    $ExistingCAPType = $AllFlexibleAssetTypes | Where-Object { $_.name -like '*Conditional Access*' } | Select-Object -First 1

    if ($ExistingCAPType) {
      $CAPTypeId = $ExistingCAPType.id
      $Message = "Found existing Conditional Access flexible asset type: $($ExistingCAPType.name) (ID: $CAPTypeId)"
    } else {
      # Create flexible asset type with all fields using relationships structure
      $NewTypeBody = @{
        data = @{
          type       = 'flexible-asset-types'
          attributes = @{
            name        = 'Conditional Access Policy'
            description = 'Microsoft 365 Conditional Access Policies synced from CIPP'
            icon        = 'shield-alt'
            enabled     = $true
          }
          relationships = @{
            'flexible-asset-fields' = @{
              data = @(
                @{
                  type       = 'flexible-asset-fields'
                  attributes = @{
                    order          = 1
                    name           = 'Policy Name'
                    kind           = 'Text'
                    required       = $true
                    'show-in-list' = $true
                  }
                }
                @{
                  type       = 'flexible-asset-fields'
                  attributes = @{
                    order          = 2
                    name           = 'Policy ID'
                    kind           = 'Text'
                    required       = $false
                    'show-in-list' = $false
                  }
                }
                @{
                  type       = 'flexible-asset-fields'
                  attributes = @{
                    order          = 3
                    name           = 'State'
                    kind           = 'Text'
                    required       = $false
                    'show-in-list' = $true
                  }
                }
                @{
                  type       = 'flexible-asset-fields'
                  attributes = @{
                    order          = 4
                    name           = 'Policy Details'
                    kind           = 'Textbox'
                    required       = $false
                    'show-in-list' = $false
                  }
                }
                @{
                  type       = 'flexible-asset-fields'
                  attributes = @{
                    order          = 5
                    name           = 'Raw JSON'
                    kind           = 'Textbox'
                    required       = $false
                    'show-in-list' = $false
                  }
                }
              )
            }
          }
        }
      } | ConvertTo-Json -Depth 20 -Compress

      $NewType = Invoke-RestMethod -Uri "$($Conn.BaseUrl)/flexible_asset_types" -Method POST -Headers $Conn.Headers -Body $NewTypeBody
      $CAPTypeId = $NewType.data.id
      $Message = "Created new Conditional Access Policy flexible asset type (ID: $CAPTypeId)"
    }

    # Save mapping to database
    $MappingTable = Get-CIPPTable -TableName CippMapping
    $AddMapping = @{
      PartitionKey    = 'ITGlueFieldMapping'
      RowKey          = 'ConditionalAccessPolicies'
      IntegrationId   = "$CAPTypeId"
      IntegrationName = 'Conditional Access Policy'
    }
    Add-CIPPAzDataTableEntity @MappingTable -Entity $AddMapping -Force

    Write-LogMessage -API $APIName -headers $Headers -message $Message -Sev 'Info'

    $Result = @{
      Success = $true
      Message = $Message
      TypeId  = $CAPTypeId
    }
    $StatusCode = [HttpStatusCode]::OK
  }
  catch {
    $ErrorMessage = Get-CippException -Exception $_
    $Result = @{
      Success = $false
      Message = "Failed to create flexible asset type: $($ErrorMessage.NormalizedError)"
    }
    Write-LogMessage -API $APIName -headers $Headers -message $Result.Message -Sev 'Error' -LogData $ErrorMessage
    $StatusCode = [HttpStatusCode]::InternalServerError
  }

  return ([HttpResponseContext]@{
      StatusCode = $StatusCode
      Body       = $Result
    })
}
