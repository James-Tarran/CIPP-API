function Format-ITGlueCAPValue {
    <#
    .FUNCTIONALITY
        Internal
    .SYNOPSIS
        Converts Out-String output (newline-separated) to comma-separated format.
        Used for formatting CAP values from Invoke-ListConditionalAccessPolicies.
    #>
    param($Value)

    if ([string]::IsNullOrWhiteSpace($Value)) { return '' }
    ($Value.Trim() -split "`n" | Where-Object { $_ -and $_.Trim() } | ForEach-Object { $_.Trim() }) -join ', '
}
