<#
    .SYNOPSIS
    Hides Silverlight updates

    .DESCRIPTION
    On some releases of Windows which have opted-in to Microsoft Update, the Silverlight application framework may be offered.

    While these updates can be hidden via the Windows Update user interface, doing so will result in the previous Silverlight update (i.e. the one which was superseded by the now hidden update) being offered.

    Hiding all Silverlight updates will typically take many "scan and hide" iterations. This script will scan for and hide all Silverlight updates in a single invocation.

    .EXAMPLE
    Hide-SilverlightUpdates

    Enumerates available updates and hides those with Silverlight in the title.

    .NOTES
    Administrator privileges are required to modify the visibility of updates.

    This is a PowerShell implementation of a WSH solution:
    https://superuser.com/a/1009947

    .LINK
    https://github.com/ralish/PSWinGlue
#>

#Requires -Version 3.0

[CmdletBinding()]
[OutputType()]
Param()

$PowerShellCore = New-Object -TypeName Version -ArgumentList 6, 0
if ($PSVersionTable.PSVersion -ge $PowerShellCore -and $PSVersionTable.Platform -ne 'Win32NT') {
    throw '{0} is only compatible with Windows.' -f $MyInvocation.MyCommand.Name
}

$User = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
if (!$User.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    throw '{0} requires Administrator privileges.' -f $MyInvocation.MyCommand.Name
}

$UpdateSession = $null
$UpdateSearcher = $null

try {
    $UpdateSession = New-Object -ComObject 'Microsoft.Update.Session'

    $UpdateSearcher = $UpdateSession.CreateUpdateSearcher()
    $UpdateSearcher.Online = $false

    do {
        $UpdatesFound = $false

        $SearchResults = $UpdateSearcher.Search('IsHidden=0 And IsInstalled=0')
        $SearchUpdates = $SearchResults.Updates

        for ($i = 0; $i -lt $SearchUpdates.Count; $i++) {
            $SearchUpdate = $SearchUpdates.Item($i)

            if ($SearchUpdate.Title -match 'Silverlight') {
                $UpdatesFound = $true

                Write-Verbose -Message ('Hiding update: {0}' -f $Update.Title)
                $Update.IsHidden = $true
            }

            $null = [Runtime.InteropServices.Marshal]::ReleaseComObject($SearchUpdate)
        }

        $null = [Runtime.InteropServices.Marshal]::ReleaseComObject($SearchUpdates)
        $null = [Runtime.InteropServices.Marshal]::ReleaseComObject($SearchResults)
    } while ($UpdatesFound)
} catch {
    throw $_
} finally {
    if ($UpdateSearcher) { $null = [Runtime.InteropServices.Marshal]::ReleaseComObject($UpdateSearcher) }
    if ($UpdateSession) { $null = [Runtime.InteropServices.Marshal]::ReleaseComObject($UpdateSession) }
}
