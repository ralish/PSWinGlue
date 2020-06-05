#Requires -Version 3.0

[CmdletBinding(SupportsShouldProcess)]
Param(
    [String[]]$Name
)

# Check required modules are present:
# - PowerShellGet: Included with PowerShell 5.0+ otherwise must be installed
$RequiredModules = @('PowerShellGet')
foreach ($Module in $RequiredModules) {
    Write-Verbose -Message ('Checking module is available: {0}' -f $Module)
    if (!(Get-Module -Name $Module -ListAvailable)) {
        throw ('Required module not available: {0}' -f $Module)
    }
}

$GetParams = @{ }
if ($PSBoundParameters.ContainsKey('Name')) {
    $GetParams['Name'] = $Name
}

Write-Verbose -Message 'Retrieving installed modules ...'
$InstalledModules = Get-InstalledModule -Verbose:$false @GetParams
Write-Verbose -Message 'Retrieving available modules ...'
$AvailableModules = Get-Module -ListAvailable -Verbose:$false @GetParams

foreach ($Module in $InstalledModules) {
    # Try to avoid subsequent call to Get-InstalledModule as it's *very* slow
    #
    # Unfortunately, we can't rely on "Get-Module -ListAvailable" due to a bug
    # in older PowerShell releases which results in modules with certain names
    # not being returned if they haven't been imported into the session.
    #
    # See: https://github.com/PowerShell/PowerShell/pull/8777
    [PSModuleInfo[]]$MatchingModules = $AvailableModules | Where-Object Name -eq $Module.Name
    if ($MatchingModules -and $MatchingModules.Count -eq 1) {
        continue
    }

    [PSCustomObject[]]$AllVersions = Get-InstalledModule -AllVersions -Name $Module.Name
    if ($AllVersions.Count -gt 1) {
        Write-Verbose -Message ('Uninstalling {0} version(s): {1}' -f $Module.Name, [String]::Join(', ', $AllVersions.Version -ne $Module.Version))
        if ($PSCmdlet.ShouldProcess($Module.Name, 'Uninstall obsolete versions')) {
            $ObsoleteModules = $AllVersions | Where-Object Version -ne $Module.Version
            foreach ($ObsoleteModule in $ObsoleteModules) {
                try {
                    $ObsoleteModule | Uninstall-Module -ErrorAction Stop
                } catch {
                    switch -Regex ($PSItem.FullyQualifiedErrorId) {
                        '^AdminPrivilegesRequiredForUninstall,' {
                            Write-Warning -Message ('Unable to uninstall module as Administrator rights are required: {0} v{1}' -f $ObsoleteModule.Name, $ObsoleteModule.Version)
                        }

                        '^UnableToUninstallAsOtherModulesNeedThisModule,' {
                            Write-Warning -Message ('Unable to uninstall module due to presence of dependent modules: {0} v{1}' -f $ObsoleteModule.Name, $ObsoleteModule.Version)
                        }

                        Default { throw }
                    }
                }
            }
        }
    }
}
