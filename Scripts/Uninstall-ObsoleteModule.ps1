#Requires -Version 3.0

[CmdletBinding(SupportsShouldProcess)]
Param(
    [String[]]$Name,

    [ValidateRange(-1, [Int]::MaxValue)]
    [Int]$ProgressParentId
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

$GetParams = @{}
if ($PSBoundParameters.ContainsKey('Name')) {
    $GetParams['Name'] = $Name
}

$WriteProgressParams = @{}

if ($PSBoundParameters.ContainsKey('ProgressParentId')) {
    $WriteProgressParams['ParentId'] = $ProgressParentId
    $WriteProgressParams['Id'] = $ProgressParentId + 1
}

if ($Name) {
    $WriteProgressParams['Activity'] = 'Uninstalling obsolete PowerShell modules (Filter: {0})' -f $Name
} else {
    $WriteProgressParams['Activity'] = 'Uninstalling obsolete PowerShell modules'
}

Write-Progress @WriteProgressParams -CurrentOperation 'Enumerating installed modules' -PercentComplete 0
$InstalledModules = Get-InstalledModule -Verbose:$false @GetParams

Write-Progress @WriteProgressParams -CurrentOperation 'Enumerating available modules' -PercentComplete 5
$AvailableModules = Get-Module -ListAvailable -Verbose:$false @GetParams

for ($ModuleIdx = 0; $ModuleIdx -lt $InstalledModules.Count; $ModuleIdx++) {
    $Module = $InstalledModules[$ModuleIdx]

    # Try to avoid subsequent call to Get-InstalledModule as it's *very* slow
    #
    # Unfortunately, we can't rely on "Get-Module -ListAvailable" due to a bug
    # in older PowerShell releases which results in modules with certain names
    # not being returned if they haven't been imported into the session.
    #
    # See: https://github.com/PowerShell/PowerShell/pull/8777
    [PSModuleInfo[]]$MatchingModules = $AvailableModules | Where-Object Name -EQ $Module.Name
    if ($MatchingModules -and $MatchingModules.Count -eq 1) {
        continue
    }

    [PSCustomObject[]]$AllVersions = Get-InstalledModule -AllVersions -Name $Module.Name
    if ($AllVersions.Count -gt 1) {
        $PercentComplete = ($ModuleIdx + 1) / $InstalledModules.Count * 90
        $CurrentOperation = 'Uninstalling {0} version(s): {1}' -f $Module.Name, [String]::Join(', ', $AllVersions.Version -ne $Module.Version)
        Write-Progress @WriteProgressParams -CurrentOperation $CurrentOperation -PercentComplete $PercentComplete

        if ($PSCmdlet.ShouldProcess($Module.Name, 'Uninstall obsolete versions')) {
            $ObsoleteModules = $AllVersions | Where-Object Version -NE $Module.Version
            foreach ($ObsoleteModule in $ObsoleteModules) {
                try {
                    $ObsoleteModule | Uninstall-Module -ErrorAction Stop
                } catch {
                    switch -Regex ($PSItem.FullyQualifiedErrorId) {
                        '^AdminPrivilegesRequiredForUninstall,' {
                            Write-Warning -Message ('Unable to uninstall module as Administrator rights are required: {0} v{1}' -f $ObsoleteModule.Name, $ObsoleteModule.Version)
                        }

                        # Uninstall-Module prints its own warning
                        '^ModuleIsInUse,' { }

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

Write-Progress @WriteProgressParams -Completed
