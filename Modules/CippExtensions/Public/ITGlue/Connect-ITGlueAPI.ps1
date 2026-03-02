function Connect-ITGlueAPI {
    <#
    .FUNCTIONALITY
        Internal
    .SYNOPSIS
        Retrieves the ITGlue API key from KeyVault and builds the connection object
        (headers + base URL) used by all subsequent ITGlue API calls.
    #>
    [CmdletBinding()]
    param (
        $Configuration
    )

    $APIKey = Get-ExtensionAPIKey -Extension 'ITGlue'

    # Resolve base URL from region setting; default to US if not set
    $Region = $Configuration.ITGlue.Region
    $BaseUrl = if ($Region -and $Region -ne '') {
        'https://{0}' -f $Region.TrimStart('https://').TrimEnd('/')
    } else {
        'https://api.itglue.com'
    }

    return [PSCustomObject]@{
        Headers = @{
            'x-api-key'    = $APIKey
            'Content-Type' = 'application/vnd.api+json'
        }
        BaseUrl = $BaseUrl
    }
}
