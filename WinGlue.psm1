Function Get-TaskSchedulerEvent {
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
