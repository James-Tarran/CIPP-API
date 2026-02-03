function Push-ExecScheduledCommand {
    <#
    .FUNCTIONALITY
        Entrypoint – executes a scheduled task or alert, safely handles multi-tenant,
        triggers, delta queries, recurrence, and ensures array output for orchestrators.
    #>
    [CmdletBinding()]
    param(
        $Item
    )

    # -----------------------------
    # Deserialize input to avoid reference issues
    # -----------------------------
    $item = $Item | ConvertTo-Json -Depth 100 | ConvertFrom-Json

    Write-Information "Running scheduled task: $($Item.TaskInfo | ConvertTo-Json -Depth 10)"

    # Async local storage for thread-safe per-invocation context
    if (-not $script:CippScheduledTaskIdStorage) {
        $script:CippScheduledTaskIdStorage = [System.Threading.AsyncLocal[string]]::new()
    }
    $script:CippScheduledTaskIdStorage.Value = $Item.TaskInfo.RowKey

    $Table = Get-CippTable -tablename 'ScheduledTasks'
    $task = $Item.TaskInfo
    $commandParameters = $Item.Parameters | ConvertTo-Json -Depth 10 | ConvertFrom-Json -AsHashtable

    $Tenant = $Item.Parameters.TenantFilter ?? $Item.TaskInfo.Tenant
    $IsMultiTenantTask = ($task.Tenant -eq 'AllTenants' -or $task.TenantGroup)
    $TenantInfo = Get-Tenants -TenantFilter $Tenant

    # -----------------------------
    # Check task existence and state
    # -----------------------------
    $CurrentTask = Get-AzDataTableEntity @Table -Filter "PartitionKey eq '$($task.PartitionKey)' and RowKey eq '$($task.RowKey)'"
    if (-not $CurrentTask) {
        Write-Information "Task $($task.Name) for tenant $Tenant does not exist. Exiting."
        Remove-Variable -Name ScheduledTaskId -Scope Script -ErrorAction SilentlyContinue
        return @()
    }

    if ($CurrentTask.TaskState -eq 'Completed' -and -not $IsMultiTenantTask) {
        Write-Information "Task $($task.Name) for tenant $Tenant is already completed. Skipping."
        Remove-Variable -Name ScheduledTaskId -Scope Script -ErrorAction SilentlyContinue
        return @()
    }

    # -----------------------------
    # Handle rerun protection
    # -----------------------------
    if ($task.Recurrence -and $task.Recurrence -ne '0') {
        $IntervalSeconds = switch -Regex ($task.Recurrence) {
            '^(\d+)$' { [int64]$matches[1] * 86400 }
            '(\d+)m$' { [int64]$matches[1] * 60 }
            '(\d+)h$' { [int64]$matches[1] * 3600 }
            '(\d+)d$' { [int64]$matches[1] * 86400 }
            default { 0 }
        }
        if ($IntervalSeconds -gt 0) {
            $FifteenMinutes = 900
            $AdjustedInterval = [Math]::Floor($IntervalSeconds / $FifteenMinutes) * $FifteenMinutes
            if ($AdjustedInterval -lt $FifteenMinutes) { $AdjustedInterval = $FifteenMinutes }
            $RerunParams = @{
                TenantFilter = $Tenant
                Type         = 'ScheduledTask'
                API          = $task.RowKey
                Interval     = $AdjustedInterval
                BaseTime     = [int64]$task.ScheduledTime
            }
            if (Test-CIPPRerun @RerunParams) {
                Write-Information "Task $($task.Name) for tenant $Tenant recently executed. Skipping."
                Remove-Variable -Name ScheduledTaskId -Scope Script -ErrorAction SilentlyContinue
                return @()
            }
        }
    }

    # -----------------------------
    # Set task state to Running
    # -----------------------------
    Update-AzDataTableEntity -Force @Table -Entity @{
        PartitionKey = $task.PartitionKey
        RowKey       = $task.RowKey
        TaskState    = 'Running'
    }

    # -----------------------------
    # Validate command exists
    # -----------------------------
    $Function = Get-Command -Name $Item.Command -ErrorAction SilentlyContinue
    if (-not $Function) {
        $Results = @("Task Failed: Command $($Item.Command) does not exist.")
        Update-AzDataTableEntity -Force @Table -Entity @{
            PartitionKey = $task.PartitionKey
            RowKey       = $task.RowKey
            Results      = ($Results | ConvertTo-Json -Compress)
            TaskState    = 'Failed'
        }
        Write-LogMessage -API 'Scheduler_UserTasks' -tenant $Tenant -tenantid $TenantInfo.customerId -message "Failed to execute task $($task.Name): Command not found" -sev Error
        Remove-Variable -Name ScheduledTaskId -Scope Script -ErrorAction SilentlyContinue
        return @($Results)
    }

    # -----------------------------
    # Filter parameters to only valid ones
    # -----------------------------
    try {
        $ValidKeys = $Function.Parameters.Keys
        foreach ($key in $commandParameters.Keys) {
            if (-not ($ValidKeys -contains $key)) {
                $commandParameters.Remove($key)
            }
        }
    } catch {
        Write-Information "Failed to filter command parameters: $($_.Exception.Message)"
    }

    # -----------------------------
    # Execute command
    # -----------------------------
    $results = @()
    try {
        Write-Information "Executing $($Item.Command) for tenant $Tenant with parameters: $($commandParameters | ConvertTo-Json -Depth 10)"
        $execResult = & $Item.Command @commandParameters
        if ($execResult -is [System.Collections.IEnumerable] -and $execResult -isnot [string]) {
            $results = $execResult
        } else {
            $results = @($execResult)
        }
    } catch {
        $results = @("Task Failed: $($_.Exception.Message)")
    }

    # -----------------------------
    # Process alerts separately
    # -----------------------------
    $TaskType = if ($Item.Command -like 'Get-CIPPAlert*') { 'Alert' } else { 'Scheduled Task' }

    try {
        if ($TaskType -ne 'Alert') {
            if ($results -is [string]) { $results = @(@{ Results = $results }) }
            $StoredResults = $results | ConvertTo-Json -Compress -Depth 20 | Out-String
        } else {
            $StoredResults = $results | ConvertTo-Json -Compress -Depth 20 | Out-String
        }

        # Write results to ScheduledTaskResults table if too large or multi-tenant
        if ($StoredResults.Length -gt 64000 -or $IsMultiTenantTask) {
            $TaskResultsTable = Get-CippTable -tablename 'ScheduledTaskResults'
            Add-CIPPAzDataTableEntity @TaskResultsTable -Entity @{
                PartitionKey = $task.RowKey
                RowKey       = $Tenant
                Results      = $StoredResults
            } -Force | Out-Null
            $StoredResults = @{ Results = 'Completed, details in More Info pane' } | ConvertTo-Json -Compress
        }

    } catch {
        Write-Warning "Failed to store task results: $($_.Exception.Message)"
    }

    # -----------------------------
    # Update task state based on recurrence
    # -----------------------------
    try {
        $secondsToAdd = switch -Regex ($task.Recurrence) {
            '(\d+)m$' { [int64]$matches[1] * 60 }
            '(\d+)h$' { [int64]$matches[1] * 3600 }
            '(\d+)d$' { [int64]$matches[1] * 86400 }
            default { 0 }
        }
        $nextRunUnixTime = [int64]$task.ScheduledTime + [int64]$secondsToAdd

        if ($task.Recurrence -eq '0' -or [string]::IsNullOrEmpty($task.Recurrence)) {
            $TaskState = 'Completed'
        } else {
            $TaskState = 'Planned'
        }

        Update-AzDataTableEntity -Force @Table -Entity @{
            PartitionKey  = $task.PartitionKey
            RowKey        = $task.RowKey
            Results       = $StoredResults
            TaskState     = $TaskState
            ScheduledTime = "$nextRunUnixTime"
        }

    } catch {
        Write-Warning "Failed to update task state: $($_.Exception.Message)"
    }

    Write-Information "Task $($task.Name) executed. TaskType: $TaskType"

    Remove-Variable -Name ScheduledTaskId -Scope Script -ErrorAction SilentlyContinue

    # -----------------------------
    # RETURN ALWAYS AS ARRAY (for AddRange)
    # -----------------------------
    if ($null -eq $results) {
        return @()
    } elseif ($results -isnot [System.Collections.IEnumerable] -or $results -is [string]) {
        return @($results)
    } else {
        return $results
    }
}
