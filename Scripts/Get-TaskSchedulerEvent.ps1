<#
    .SYNOPSIS
    Retrieves events matching the specified IDs from the Task Scheduler event log

    .DESCRIPTION
    Retrieves specified event IDs from the Task Scheduler event log, up to a maximum number of events, with optional filtering of specific task names.

    .PARAMETER EventIds
    The event IDs to query for.

    The default is to query for event IDs which represent a scheduled task failure: 111, 202, 203, 323, 329, 331.

    .PARAMETER IgnoredTasks
    An optional array of task names to filter out of the returned results.

    The default is to exclude several tasks for which failure is typically benign:
    - \Microsoft\Windows\.NET Framework\.NET Framework NGEN v4.0.30319
    - \Microsoft\Windows\.NET Framework\.NET Framework NGEN v4.0.30319 64
    - \Microsoft\Windows\NetCfg\BindingWorkItemQueueHandler
    - \Microsoft\Windows\Shell\CreateObjectTas

    Note that filtering of ignored tasks is performed after the specified maximum number of events have been returned.

    .PARAMETER MaxEvents
    Maximum number of events to return.

    The default is 100 events.

    Note that this parameter interacts with the IgnoredTasks parameter in a way that may be counter-intuitive.

    .EXAMPLE
    Get-TaskSchedulerEvent

    Retrieves the most recent 100 events indicating a scheduled task failure, ignoring typically benign task failures.

    .NOTES
    Task Scheduler events
    https://learn.microsoft.com/en-us/previous-versions/windows/it-pro/windows-server-2008-R2-and-2008/dd315533(v%3dws.10)#events

    .LINK
    https://github.com/ralish/PSWinGlue
#>

#Requires -Version 3.0

[CmdletBinding()]
[OutputType([Diagnostics.Eventing.Reader.EventRecord[]])]
Param(
    [ValidateNotNullOrEmpty()]
    [Int[]]$EventIds = @(111, 202, 203, 323, 329, 331),

    [ValidateRange(1, 1000)]
    [Int]$MaxEvents = 100,

    [String[]]$IgnoredTasks = @(
        '\Microsoft\Windows\.NET Framework\.NET Framework NGEN v4.0.30319',
        '\Microsoft\Windows\.NET Framework\.NET Framework NGEN v4.0.30319 64',
        '\Microsoft\Windows\NetCfg\BindingWorkItemQueueHandler',
        '\Microsoft\Windows\Shell\CreateObjectTask'
    )
)

$PowerShellCore = New-Object -TypeName Version -ArgumentList 6, 0
if ($PSVersionTable.PSVersion -ge $PowerShellCore -and $PSVersionTable.Platform -ne 'Win32NT') {
    throw '{0} is only compatible with Windows.' -f $MyInvocation.MyCommand.Name
}

$WinEventParams = @{
    FilterHashtable = @{
        ProviderName = 'Microsoft-Windows-TaskScheduler'
        ID           = $EventIds
    }

    MaxEvents       = $MaxEvents
    ErrorAction     = 'Stop'
}

$Events = [Collections.Generic.List[Diagnostics.Eventing.Reader.EventRecord]]@(Get-WinEvent @WinEventParams)

if ($IgnoredTasks) {
    $FilteredEvents = New-Object -TypeName 'Collections.Generic.List[Diagnostics.Eventing.Reader.EventRecord]'

    foreach ($Event in $Events) {
        $EventXml = [Xml]$Event.ToXml()
        $EventData = $EventXml.Event.EventData.Data
        $TaskName = $EventData | Where-Object Name -EQ 'TaskName'

        if ($TaskName -notin $IgnoredTasks) {
            $FilteredEvents.Add($Event)
        }
    }

    $Events = $FilteredEvents
}

return $Events.ToArray()
