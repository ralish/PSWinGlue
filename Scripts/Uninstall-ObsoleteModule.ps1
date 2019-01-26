[CmdletBinding(SupportsShouldProcess)]
Param(
    [String[]]$Name
)

$GetParams = @{}
if ($PSBoundParameters.ContainsKey('Name')) {
    $GetParams['Name'] = $Name
}

Write-Verbose -Message 'Retrieving installed modules ...'
$InstalledModules = Get-InstalledModule -Verbose:$false @GetParams
Write-Verbose -Message 'Retrieving available modules ...'
$AvailableModules = Get-Module -ListAvailable -Verbose:$false @GetParams

foreach ($Module in $InstalledModules) {
    # Try to avoid subsequent call to Get-InstalledModule as it's obscenely slow
    [PSModuleInfo[]]$MatchingModules = $AvailableModules | Where-Object Name -eq $Module.Name
    if ($MatchingModules.Count -eq 1) {
        continue
    }

    $AllVersions = Get-InstalledModule -AllVersions -Name $Module.Name
    if ($AllVersions.Count -gt 1) {
        Write-Verbose -Message ('Uninstalling {0} version(s): {1}' -f $Module.Name, [String]::Join(', ', $AllVersions.Version -ne $Module.Version))
        if ($PSCmdlet.ShouldProcess($Module.Name, 'Uninstall obsolete versions')) {
            $AllVersions | Where-Object Version -ne $Module.Version | Uninstall-Module
        }
    }
}
