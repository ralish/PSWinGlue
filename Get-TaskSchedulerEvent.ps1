Function Get-TaskSchedulerEvent {
    <#
        .Synopsis
        Fetches events matching given IDs from the Task Scheduler event log.
        .Parameter EventIds
        An array of integers specifying the event IDs we want to match in the query. The default is
        "(111,202,203,323,329,331)" which corresponds to all event IDs which represent a failure to
        complete a scheduled task. This is particularly useful for identifying task failures.
        .Parameter MaxEvents
        Specifies the maximum number of events to return. Note that any filtering of returned events
        by the IgnoredTasks parameter is performed after the specified maximum number of events have
        been returned. As such, you cannot rely on receiving the maximum number of results as set by
        MaxEvents, even if there are enough events in the Event Log with filtering applied.
        .Parameter IgnoredTasks
        An array of strings specifying the task names we wish to exclude from the returned results.
        .Notes
        For more information on Task Scheduler event IDs consult the TechNet documentation at:
        http://technet.microsoft.com/en-us/library/dd363729%28v=ws.10%29.aspx

    #>
    Param(
        [Int32[]]$EventIds=(111,202,203,323,329,331),
        [Int32]$MaxEvents=10,
        [String[]]$IgnoredTasks=('\Microsoft\Windows\Shell\CreateObjectTask')
    )

    $ErrorActionPreference = 'Stop'
    $LogName = 'Microsoft-Windows-TaskScheduler/Operational'

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
