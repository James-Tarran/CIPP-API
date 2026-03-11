function New-ITGlueCAPolicyAssetType {
    <#
    .FUNCTIONALITY
        Internal
    .SYNOPSIS
        Creates the Conditional Access Policy Flexible Asset Type in ITGlue with all required fields.
    .DESCRIPTION
        Creates a new flexible asset type in ITGlue with 50+ fields to store detailed
        conditional access policy information from Microsoft 365.
    #>
    [CmdletBinding()]
    param()

    $Table = Get-CIPPTable -TableName Extensionsconfig
    try {
        $Configuration = (Get-CIPPAzDataTableEntity @Table).config | ConvertFrom-Json -ErrorAction Stop
        $Conn = Connect-ITGlueAPI -Configuration $Configuration

        # Define all the flexible asset fields for CA Policies
        # Field positions are 1-based and sequential
        $Fields = @(
            # Policy name (required, use for title)
            @{ name = 'Policy name'; kind = 'Text'; required = $true; 'use-for-title' = $true; 'show-in-list' = $true }
            @{ name = 'Policy description'; kind = 'Text'; required = $false; 'show-in-list' = $false }
            @{ name = 'Policy state'; kind = 'Select'; required = $false; 'show-in-list' = $true; 'default-value' = 'enabled,enabledForReportingButNotEnforced,disabled' }

            # Users section
            @{ name = 'Users'; kind = 'Header' }
            @{ name = 'Users include'; kind = 'Select'; required = $false; 'default-value' = 'All users,None,Select users and groups,All guest and external users' }
            @{ name = 'Included guest or external roles'; kind = 'Textbox'; required = $false }
            @{ name = 'Included directory roles'; kind = 'Textbox'; required = $false }
            @{ name = 'Included users and groups'; kind = 'Textbox'; required = $false }
            @{ name = 'Users exclude'; kind = 'Select'; required = $false; 'default-value' = 'None,Select users and groups,All guest and external users' }
            @{ name = 'Excluded guest or external roles'; kind = 'Textbox'; required = $false }
            @{ name = 'Excluded directory roles'; kind = 'Textbox'; required = $false }
            @{ name = 'Excluded users and groups'; kind = 'Textbox'; required = $false }

            # Target resources section
            @{ name = 'Target resources'; kind = 'Header' }
            @{ name = 'Target resources include'; kind = 'Select'; required = $false; 'default-value' = 'All cloud apps,None,Select apps,User actions,Authentication context' }
            @{ name = 'Select apps'; kind = 'Textbox'; required = $false }
            @{ name = 'User actions'; kind = 'Textbox'; required = $false }
            @{ name = 'Authentication context'; kind = 'Textbox'; required = $false }
            @{ name = 'Target resources excluded cloud apps'; kind = 'Textbox'; required = $false }

            # Network section
            @{ name = 'Network'; kind = 'Header' }
            @{ name = 'Network include'; kind = 'Select'; required = $false; 'default-value' = 'Any network or location,All trusted networks and locations,Selected networks and locations' }
            @{ name = 'Network include - selected networks and locations'; kind = 'Textbox'; required = $false }
            @{ name = 'Network exclude'; kind = 'Select'; required = $false; 'default-value' = 'None,All trusted networks and locations,Selected networks and locations' }
            @{ name = 'Network exclude - selected networks and locations'; kind = 'Textbox'; required = $false }

            # Conditions section
            @{ name = 'Conditions'; kind = 'Header' }

            # User risk
            @{ name = 'User risk - high'; kind = 'Checkbox'; required = $false }
            @{ name = 'User risk - medium'; kind = 'Checkbox'; required = $false }
            @{ name = 'User risk - low'; kind = 'Checkbox'; required = $false }

            # Sign-in risk
            @{ name = 'Sign-in risk - high'; kind = 'Checkbox'; required = $false }
            @{ name = 'Sign-in risk - medium'; kind = 'Checkbox'; required = $false }
            @{ name = 'Sign-in risk - low'; kind = 'Checkbox'; required = $false }
            @{ name = 'Sign-in risk - no risk'; kind = 'Checkbox'; required = $false }

            # Insider risk
            @{ name = 'Insider risk - elevated'; kind = 'Checkbox'; required = $false }
            @{ name = 'Insider risk - moderate'; kind = 'Checkbox'; required = $false }
            @{ name = 'Insider risk - minor'; kind = 'Checkbox'; required = $false }

            # Device platforms - Include
            @{ name = 'Device platforms'; kind = 'Header' }
            @{ name = 'Include Android'; kind = 'Checkbox'; required = $false }
            @{ name = 'Include iOS'; kind = 'Checkbox'; required = $false }
            @{ name = 'Include Windows'; kind = 'Checkbox'; required = $false }
            @{ name = 'Include macOS'; kind = 'Checkbox'; required = $false }
            @{ name = 'Include Linux'; kind = 'Checkbox'; required = $false }
            @{ name = 'Include Windows Phone'; kind = 'Checkbox'; required = $false }

            # Device platforms - Exclude
            @{ name = 'Exclude Android'; kind = 'Checkbox'; required = $false }
            @{ name = 'Exclude iOS'; kind = 'Checkbox'; required = $false }
            @{ name = 'Exclude Windows'; kind = 'Checkbox'; required = $false }
            @{ name = 'Exclude macOS'; kind = 'Checkbox'; required = $false }
            @{ name = 'Exclude Linux'; kind = 'Checkbox'; required = $false }
            @{ name = 'Exclude Windows Phone'; kind = 'Checkbox'; required = $false }

            # Client apps
            @{ name = 'Client apps'; kind = 'Header' }
            @{ name = 'Client apps configured'; kind = 'Checkbox'; required = $false }
            @{ name = 'Browser'; kind = 'Checkbox'; required = $false }
            @{ name = 'Mobile apps and desktop clients'; kind = 'Checkbox'; required = $false }
            @{ name = 'Exchange ActiveSync clients'; kind = 'Checkbox'; required = $false }
            @{ name = 'Other clients'; kind = 'Checkbox'; required = $false }

            # Device filters
            @{ name = 'Filter for devices'; kind = 'Header' }
            @{ name = 'Device filter mode'; kind = 'Select'; required = $false; 'default-value' = 'Not configured,Include,Exclude' }
            @{ name = 'Device filter rule'; kind = 'Textbox'; required = $false }

            # Grant controls
            @{ name = 'Grant controls'; kind = 'Header' }
            @{ name = 'Grant or Block'; kind = 'Select'; required = $false; 'default-value' = 'Grant access,Block access' }
            @{ name = 'Grant controls operator'; kind = 'Select'; required = $false; 'default-value' = 'AND,OR' }
            @{ name = 'Require multifactor authentication'; kind = 'Checkbox'; required = $false }
            @{ name = 'Require authentication strength'; kind = 'Text'; required = $false }
            @{ name = 'Require device to be marked as compliant'; kind = 'Checkbox'; required = $false }
            @{ name = 'Require Microsoft Entra hybrid joined device'; kind = 'Checkbox'; required = $false }
            @{ name = 'Require approved client app'; kind = 'Checkbox'; required = $false }
            @{ name = 'Require app protection policy'; kind = 'Checkbox'; required = $false }
            @{ name = 'Require password change'; kind = 'Checkbox'; required = $false }
            @{ name = 'Terms of use'; kind = 'Textbox'; required = $false }

            # Session controls
            @{ name = 'Session controls'; kind = 'Header' }
            @{ name = 'Use app enforced restrictions'; kind = 'Checkbox'; required = $false }
            @{ name = 'Use Conditional Access App Control'; kind = 'Checkbox'; required = $false }
            @{ name = 'Conditional Access App Control type'; kind = 'Text'; required = $false }
            @{ name = 'Sign-in frequency enabled'; kind = 'Checkbox'; required = $false }
            @{ name = 'Sign-in frequency value'; kind = 'Text'; required = $false }
            @{ name = 'Sign-in frequency type'; kind = 'Select'; required = $false; 'default-value' = 'hours,days,everyTime' }
            @{ name = 'Persistent browser session enabled'; kind = 'Checkbox'; required = $false }
            @{ name = 'Persistent browser session mode'; kind = 'Select'; required = $false; 'default-value' = 'always,never' }
            @{ name = 'Continuous access evaluation'; kind = 'Select'; required = $false; 'default-value' = 'Not configured,Disabled,Strictly enforced' }
            @{ name = 'Disable resilience defaults'; kind = 'Checkbox'; required = $false }
            @{ name = 'Secure sign-in session'; kind = 'Checkbox'; required = $false }

            # Metadata
            @{ name = 'Metadata'; kind = 'Header' }
            @{ name = 'Policy ID'; kind = 'Text'; required = $false }
            @{ name = 'Created date'; kind = 'Text'; required = $false }
            @{ name = 'Modified date'; kind = 'Text'; required = $false }
            @{ name = 'CIPP link'; kind = 'Text'; required = $false }
            @{ name = 'Entra link'; kind = 'Text'; required = $false }
            @{ name = 'Last synced'; kind = 'Text'; required = $false }
        )

        # Add position to each field
        $Position = 1
        $FieldsWithPosition = foreach ($Field in $Fields) {
            $Field['position'] = $Position
            $Position++
            $Field
        }

        # Create the flexible asset type
        $TypeBody = @{
            data = @{
                type       = 'flexible-asset-types'
                attributes = @{
                    name                      = 'M365 Conditional Access Policy'
                    description               = 'Microsoft 365 Conditional Access Policies synced from CIPP'
                    icon                      = 'shield-check'
                    enabled                   = $true
                    'flexible-asset-fields'   = @($FieldsWithPosition)
                }
            }
        } | ConvertTo-Json -Depth 20 -Compress

        $Response = Invoke-RestMethod -Uri "$($Conn.BaseUrl)/flexible_asset_types" -Method POST -Headers $Conn.Headers -Body $TypeBody

        $CreatedTypeId = $Response.data.id
        $CreatedTypeName = $Response.data.attributes.name

        Write-LogMessage -Message "Created ITGlue Conditional Access Policy flexible asset type: $CreatedTypeName (ID: $CreatedTypeId)" -Level Info -tenant 'CIPP' -API 'ITGlueMapping'

        return @{
            Success = $true
            Message = "Successfully created flexible asset type '$CreatedTypeName' (ID: $CreatedTypeId). You can now select it in the field mapping dropdown."
            TypeId  = $CreatedTypeId
            TypeName = $CreatedTypeName
        }

    } catch {
        $ErrorMessage = if ($_.ErrorDetails.Message) {
            try {
                $ErrorBody = $_.ErrorDetails.Message | ConvertFrom-Json -ErrorAction SilentlyContinue
                $ErrorBody.errors[0].detail ?? $ErrorBody.errors[0].title ?? $_.ErrorDetails.Message
            } catch {
                $_.ErrorDetails.Message
            }
        } else {
            $_.Exception.Message
        }

        Write-LogMessage -Message "Failed to create ITGlue CA Policy flexible asset type: $ErrorMessage" -Level Error -tenant 'CIPP' -API 'ITGlueMapping'

        return @{
            Success = $false
            Message = "Failed to create flexible asset type: $ErrorMessage"
        }
    }
}
