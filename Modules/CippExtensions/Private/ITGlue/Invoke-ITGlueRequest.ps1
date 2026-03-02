function Invoke-ITGlueRequest {
    <#
    .FUNCTIONALITY
        Internal
    .SYNOPSIS
        Core HTTP helper for all ITGlue API calls. Handles JSON:API wrapping/unwrapping,
        pagination, and rate limiting. No third-party module required.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('GET', 'POST', 'PATCH', 'DELETE')]
        [string]$Method,

        [Parameter(Mandatory = $true)]
        [string]$Endpoint,

        [Parameter(Mandatory = $true)]
        [hashtable]$Headers,

        [Parameter(Mandatory = $true)]
        [string]$BaseUrl,

        # For POST/PATCH: the attributes hashtable (will be wrapped in JSON:API envelope)
        [Parameter(Mandatory = $false)]
        [hashtable]$Attributes,

        # The JSON:API resource type (e.g. 'flexible-assets', 'contacts', 'configurations')
        [Parameter(Mandatory = $false)]
        [string]$ResourceType,

        # For PATCH: the ID of the resource to update
        [Parameter(Mandatory = $false)]
        [string]$ResourceId,

        # Query string parameters as a hashtable (e.g. @{ 'filter[organization_id]' = 123 })
        [Parameter(Mandatory = $false)]
        [hashtable]$QueryParams,

        # Page size for GET requests (max 1000, default 50)
        [Parameter(Mandatory = $false)]
        [int]$PageSize = 50,

        # If set, only retrieve the first page (useful for connection tests)
        [Parameter(Mandatory = $false)]
        [switch]$FirstPageOnly
    )

    $Uri = '{0}{1}' -f $BaseUrl.TrimEnd('/'), $Endpoint

    # Build query string
    $AllQueryParams = @{}
    if ($QueryParams) {
        foreach ($Key in $QueryParams.Keys) {
            $AllQueryParams[$Key] = $QueryParams[$Key]
        }
    }

    # Build JSON:API request body for POST/PATCH
    $Body = $null
    if ($Method -in @('POST', 'PATCH') -and $Attributes) {
        $DataObject = @{
            type       = $ResourceType
            attributes = $Attributes
        }
        if ($ResourceId) {
            $DataObject['id'] = $ResourceId
        }
        $Body = @{ data = $DataObject } | ConvertTo-Json -Depth 20 -Compress
    }

    $Results = [System.Collections.Generic.List[object]]::new()

    $CurrentPage = 1
    do {
        $PageQueryParams = $AllQueryParams.Clone()
        if ($Method -eq 'GET') {
            $PageQueryParams['page[size]']   = $PageSize
            $PageQueryParams['page[number]'] = $CurrentPage
        }

        # Build URI with query string
        $QueryString = ($PageQueryParams.GetEnumerator() | ForEach-Object {
            '{0}={1}' -f [System.Uri]::EscapeDataString($_.Key), [System.Uri]::EscapeDataString($_.Value)
        }) -join '&'
        $FullUri = if ($QueryString) { '{0}?{1}' -f $Uri, $QueryString } else { $Uri }

        $InvokeParams = @{
            Uri     = $FullUri
            Method  = $Method
            Headers = $Headers
        }
        if ($Body) {
            $InvokeParams['Body'] = $Body
        }

        try {
            $Response = Invoke-RestMethod @InvokeParams
        } catch {
            $StatusCode = $_.Exception.Response.StatusCode.value__
            $ErrorDetail = $null
            try {
                $ErrorBody = $_.ErrorDetails.Message | ConvertFrom-Json -ErrorAction SilentlyContinue
                $ErrorDetail = $ErrorBody.errors[0].detail ?? $ErrorBody.errors[0].title ?? $_.ErrorDetails.Message
            } catch {}
            $ErrorMessage = if ($ErrorDetail) { $ErrorDetail } else { $_.Exception.Message }
            throw "ITGlue API error ($StatusCode) on $Method $Endpoint : $ErrorMessage"
        }

        # Unwrap JSON:API response — flatten attributes onto a new object and add the top-level id
        if ($Response.data) {
            $DataItems = if ($Response.data -is [array]) { $Response.data } else { @($Response.data) }
            foreach ($Item in $DataItems) {
                # Select-Object * copies all NoteProperty members into a fresh PSCustomObject
                $Obj = $Item.attributes | Select-Object *
                $Obj | Add-Member -NotePropertyName 'id' -NotePropertyValue $Item.id -Force
                $Results.Add($Obj)
            }
        }

        # Pagination — only continue for GET requests collecting all pages
        if ($Method -ne 'GET' -or $FirstPageOnly) { break }

        $TotalPages = $Response.meta.'total-pages'
        if (-not $TotalPages -or $CurrentPage -ge $TotalPages) { break }

        $CurrentPage++
        # Respect ITGlue rate limits between paginated calls
        Start-Sleep -Milliseconds 250

    } while ($true)

    return $Results.ToArray()
}
