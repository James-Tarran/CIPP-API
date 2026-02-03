function Write-AlertTrace {
    <#
    .FUNCTIONALITY
        Internal function. Writes alert trace data and guarantees
        a consistent array return type for schedulers.
    #>
    [CmdletBinding()]
    param(
        $cmdletName,
        $data,
        $tenantFilter,
        [string]$PartitionKey = (Get-Date -UFormat '%Y%m%d').ToString(),
        [string]$AlertComment = $null
    )

    # -----------------------------
    # Normalize input data EARLY
    # -----------------------------
    if ($null -eq $data) {
        $data = @()
    }
    elseif ($data -is [string] -or $data -isnot [System.Collections.IEnumerable]) {
        $data = @($data)
    }

    $Table = Get-CIPPTable -tablename AlertLastRun

    try {
        # Get existing row (if any)
        $Row = Get-CIPPAzDataTableEntity @Table -Filter "RowKey eq '$($tenantFilter)-$($cmdletName)' and PartitionKey eq '$PartitionKey'"

        $CurrentJson = ConvertTo-Json -InputObject $data -Compress -Depth 10 | Out-String
        $PreviousJson = $Row.LogData

        # Only write if data changed
        if ($PreviousJson -ne $CurrentJson) {
            $TableRow = @{
                PartitionKey = $PartitionKey
                RowKey       = "$($tenantFilter)-$($cmdletName)"
                CmdletName   = "$cmdletName"
                Tenant       = "$tenantFilter"
                LogData      = [string]$CurrentJson
                AlertComment = [string]$AlertComment
            }

            $Table.Entity = $TableRow
            Add-CIPPAzDataTableEntity @Table -Force | Out-Null
        }

    } catch {
        # First run or lookup failure — always write
        $CurrentJson = ConvertTo-Json -InputObject $data -Compress -Depth 10 | Out-String

        $TableRow = @{
            PartitionKey = $PartitionKey
            RowKey       = "$($tenantFilter)-$($cmdletName)"
            CmdletName   = "$cmdletName"
            Tenant       = "$tenantFilter"
            LogData      = [string]$CurrentJson
            AlertComment = [string]$AlertComment
        }

        $Table.Entity = $TableRow
        Add-CIPPAzDataTableEntity @Table -Force | Out-Null
    }

    # -----------------------------
    # ALWAYS return an array
    # -----------------------------
    return $data
}
