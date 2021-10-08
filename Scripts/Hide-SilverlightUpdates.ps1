#Requires -Version 3.0
#Requires -RunAsAdministrator

[CmdletBinding()]
Param()

# PowerShell implementation of WSH solution:
# https://superuser.com/a/1009947

try {
    $UpdateSession = New-Object -ComObject Microsoft.Update.Session
} catch {
    throw $_
}

try {
    $UpdateSearcher = $UpdateSession.CreateUpdateSearcher()
    $UpdateSearcher.Online = $false
} catch {
    throw $_
}

do {
    $UpdatesFound = $false
    $SearchResults = $UpdateSearcher.Search("IsHidden=0 And IsInstalled=0")

    if ($SearchResults.Updates.Count -gt 0) {
        foreach ($Update in ($SearchResults.Updates | Where-Object Title -Match 'Silverlight')) {
            Write-Verbose -Message ('Hiding update: {0}' -f $Update.Title)
            $Update.IsHidden = $true
            $UpdatesFound = $true
        }
    }

    $null = [Runtime.InteropServices.Marshal]::FinalReleaseComObject($SearchResults)
} while ($UpdatesFound)

$null = [Runtime.InteropServices.Marshal]::FinalReleaseComObject($UpdateSearcher)
$null = [Runtime.InteropServices.Marshal]::FinalReleaseComObject($UpdateSession)
