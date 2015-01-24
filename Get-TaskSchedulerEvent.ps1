Function Get-TaskSchedulerEvent {
    <#
        .Synopsis
        Fetches events matching given IDs from the Task Scheduler event log.
        .Description
        .Parameter EventIds
        An array of integers specifying the event IDs we want to match in the query. A particularly
        useful set is "(111,202,203,323,329,331)", which corresponds to all current event IDs which
        represent some form of failure to complete a scheduled task.
        See: http://technet.microsoft.com/en-us/library/dd363729%28v=ws.10%29.aspx
        .Parameter MaxEvents
        Specifies the maximum number of events to return. Note that any filtering of returned events
        by the IgnoredTasks parameter is performed after the specified maximum number of events have
        been returned. As such, you cannot rely on receiving the maximum number of results as set by
        MaxEvents, even if there are enough events in the Event Log with filtering applied.
        .Parameter LogName
        The Event Log name to query; the default should rarely if ever need to be changed.
        .Parameter IgnoredTasks
        An array of strings specifying the task names we wish to exclude from the returned results.
        .Notes
    #>
    Param(
        [Parameter(Mandatory=$true)]
            [Int32[]]$EventIds,
        [Int32]$MaxEvents=1,
        [String]$LogName='Microsoft-Windows-TaskScheduler/Operational',
        [String[]]$IgnoredTasks=('\Microsoft\Windows\Shell\CreateObjectTask')
    )

    $ErrorActionPreference = 'Stop'

    # Fetch the most recent matching scheduled task event
    Write-Verbose 'Fetching matching task scheduler events...'
    $Event = Get-WinEvent -FilterHashTable @{ LogName = $LogName; ID = $EventIds } -MaxEvents $MaxEvents

    # Exit if the event matches any defined ignored tasks
    Write-Verbose 'Filtering out any ignored task events...'
    $EventXml = [xml]$Event.ToXml()
    $EventTaskName = $EventXml.Event.EventData.Data | ? {$_.Name -eq 'TaskName'}
    if ($EventTaskName.'#text' -in $IgnoredTasks) {
        return
    } else {
        return $Event
    }
}
