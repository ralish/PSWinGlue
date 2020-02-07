<#
    .SYNOPSIS
    Fetches events matching given IDs from the Task Scheduler event log

    .DESCRIPTION
    Alternate data streams can contain important (meta)data which shouldn't be removed, however, they can also be a nuisance.

    This is particularly true of certain common alternate data streams used to perform additional prompting of the user for "untrusted" files.

    This cmdlet provides options to remove common alternate data streams which may be unwanted while preserving any other present alternate data streams.

    .PARAMETER EventIds
    An array of integers specifying the event IDs we want to match in the query. The default is "(111,202,203,323,329,331)" which corresponds to all event IDs which represent a failure to complete a scheduled task. This is particularly useful for identifying task failures.

    .PARAMETER IgnoredTasks
    An array of strings specifying the task names we wish to exclude from the returned results.

    .PARAMETER MaxEvents
    Specifies the maximum number of events to return. Note that any filtering of returned events by the IgnoredTasks parameter is performed after the specified maximum number of events have been returned. As such, you cannot rely on receiving the maximum number of results as set by MaxEvents, even if there are enough events in the Event Log with filtering applied.

    .EXAMPLE
    Remove-AlternateDataStreams -Path D:\Library -ZoneIdentifier -Recurse

    Removes all Zone.Identifier alternate data streams recursively from files in D:\Library.

    .NOTES
    For more information on Task Scheduler event IDs consult the TechNet documentation at:
    https://docs.microsoft.com/en-us/previous-versions/windows/it-pro/windows-server-2008-R2-and-2008/dd315533(v%3dws.10)

    .LINK
    https://github.com/ralish/PSWinGlue
#>

#Requires -Version 3.0

[CmdletBinding()]
Param(
    [Int[]]$EventIds=(111, 202, 203, 323, 329, 331),
    [Int]$MaxEvents=10,
    [String[]]$IgnoredTasks=(
        '\Microsoft\Windows\.NET Framework\.NET Framework NGEN v4.0.30319',
        '\Microsoft\Windows\.NET Framework\.NET Framework NGEN v4.0.30319 64',
        '\Microsoft\Windows\NetCfg\BindingWorkItemQueueHandler',
        '\Microsoft\Windows\Shell\CreateObjectTask'
    )
)

$Events = @()
$Events += Get-WinEvent -FilterHashTable @{ ProviderName = 'Microsoft-Windows-TaskScheduler'; ID = $EventIds } -MaxEvents $MaxEvents

if ($IgnoredTasks) {
    $FilteredEvents = @()

    foreach ($Event in $Events) {
        $EventXml = [Xml]$Event.ToXml()
        $EventData = $EventXml.Event.EventData.Data
        $TaskName = $EventData | Where-Object Name -eq 'TaskName'

        if ($TaskName -notin $IgnoredTasks) {
            $FilteredEvents += $Event
        }
    }

    $Events = $FilteredEvents
}

if (!$Events) {
    Write-Warning -Message 'No events returned for the given filter.'
}

return $Events
