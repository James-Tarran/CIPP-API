function New-HaloPSATicket {
  [CmdletBinding(SupportsShouldProcess)]
  param (
    $Title,
    $Description,
    $Client,
    $TicketId
  )

  # Load Halo configuration
  $Table = Get-CIPPTable -TableName Extensionsconfig
  $Configuration = ((Get-CIPPAzDataTableEntity @Table).config | ConvertFrom-Json).HaloPSA
  $TicketTable = Get-CIPPTable -TableName 'PSATickets'
  $Token = Get-HaloToken -configuration $Configuration

  # Helper to add a note to an existing ticket
  function Add-HaloTicketNote {
    param ($TicketId, $Html)

    $Object = [PSCustomObject]@{
      ticket_id      = $TicketId
      outcome_id     = 7
      hiddenfromuser = $true
      note_html      = $Html
    }

    if ($Configuration.Outcome) {
      $Object.outcome_id = $Configuration.Outcome.value ?? $Configuration.Outcome
    }

    $Body = ConvertTo-Json -Compress -Depth 10 -InputObject @($Object)

    if ($PSCmdlet.ShouldProcess("HaloPSA Ticket $TicketId", 'Add note')) {
      Invoke-RestMethod `
        -Uri "$($Configuration.ResourceURL)/actions" `
        -ContentType 'application/json; charset=utf-8' `
        -Method Post `
        -Body $Body `
        -Headers @{ Authorization = "Bearer $($Token.access_token)" }
    }
  }

  if ($TicketId) {
    Write-Information "Explicit PSA Ticket ID provided: $TicketId"

    try {
      $Ticket = Invoke-RestMethod `
        -Uri "$($Configuration.ResourceURL)/Tickets/$TicketId?includedetails=true&includelastaction=false" `
        -ContentType 'application/json; charset=utf-8' `
        -Method Get `
        -Headers @{ Authorization = "Bearer $($Token.access_token)" } `
        -SkipHttpErrorCheck

      if ($Ticket.id -and -not $Ticket.hasbeenclosed) {
        Write-Information "Ticket $TicketId is open. Appending note."
        Add-HaloTicketNote -TicketId $TicketId -Html $Description
        return "Note added to HaloPSA ticket $TicketId"
      }

      Write-Information "Ticket $TicketId is closed or not found. Creating new ticket."
    }
    catch {
      $Message = $_.Exception.Message
      Write-LogMessage `
        -API 'HaloPSATicket' `
        -sev Error `
        -message "Failed to update HaloPSA ticket $($TicketId): $Message" `
        -LogData (Get-CippException -Exception $_)
      return "Failed to update HaloPSA ticket $($TicketId): $Message"
    }
  }

  $TitleHash = Get-StringHash -String $Title

  if ($Configuration.ConsolidateTickets) {
    $ExistingTicket = Get-CIPPAzDataTableEntity `
      @TicketTable `
      -Filter "PartitionKey eq 'HaloPSA' and RowKey eq '$($Client)-$($TitleHash)'"

    if ($ExistingTicket) {
      Write-Information "Consolidated ticket found: $($ExistingTicket.TicketID)"

      try {
        $Ticket = Invoke-RestMethod `
          -Uri "$($Configuration.ResourceURL)/Tickets/$($ExistingTicket.TicketID)?includedetails=true&includelastaction=false" `
          -ContentType 'application/json; charset=utf-8' `
          -Method Get `
          -Headers @{ Authorization = "Bearer $($Token.access_token)" } `
          -SkipHttpErrorCheck

        if ($Ticket.id -and -not $Ticket.hasbeenclosed) {
          Write-Information "Consolidated ticket open. Appending note."
          Add-HaloTicketNote -TicketId $ExistingTicket.TicketID -Html $Description
          return "Note added to HaloPSA ticket $($ExistingTicket.TicketID)"
        }

        Write-Information "Consolidated ticket closed. Creating new ticket."
      }
      catch {
        Write-Information "Failed to read consolidated ticket. Creating new ticket."
      }
    }
  }

  $Object = [PSCustomObject]@{
    files                      = $null
    usertype                   = 1
    userlookup                 = @{
      id            = -1
      lookupdisplay = 'Enter Details Manually'
    }
    client_id                  = ($Client | Select-Object -Last 1)
    _forcereassign             = $true
    site_id                    = $null
    user_name                  = $null
    reportedby                 = $null
    summary                    = $Title
    details_html               = $Description
    donotapplytemplateintheapi = $true
    attachments                = @()
    _novalidate                = $true
  }

  if ($Configuration.TicketType) {
    $TicketType = $Configuration.TicketType.value ?? $Configuration.TicketType
    $Object | Add-Member -MemberType NoteProperty -Name 'tickettype_id' -Value $TicketType -Force
  }

  $Body = ConvertTo-Json -Compress -Depth 10 -InputObject @($Object)

  Write-Information 'Creating new HaloPSA ticket'

  try {
    if ($PSCmdlet.ShouldProcess('HaloPSA', 'Create ticket')) {
      $Ticket = Invoke-RestMethod `
        -Uri "$($Configuration.ResourceURL)/Tickets" `
        -ContentType 'application/json; charset=utf-8' `
        -Method Post `
        -Body $Body `
        -Headers @{ Authorization = "Bearer $($Token.access_token)" }

      Write-Information "Ticket created in HaloPSA: $($Ticket.id)"

      if ($Configuration.ConsolidateTickets) {
        $TicketObject = [PSCustomObject]@{
          PartitionKey = 'HaloPSA'
          RowKey       = "$($Client)-$($TitleHash)"
          Title        = $Title
          ClientId     = $Client
          TicketID     = $Ticket.id
        }
        Add-CIPPAzDataTableEntity @TicketTable -Entity $TicketObject -Force
        Write-Information 'Ticket added to consolidation table'
      }

      return "Ticket created in HaloPSA: $($Ticket.id)"
    }
  }
  catch {
    $Message = $_.Exception.Message
    Write-LogMessage `
      -API 'HaloPSATicket' `
      -sev Error `
      -message "Failed to create HaloPSA ticket: $Message" `
      -LogData (Get-CippException -Exception $_)
    return "Failed to create HaloPSA ticket: $Message"
  }
}
