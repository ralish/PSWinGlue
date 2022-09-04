<#
    .SYNOPSIS
    Uninstalls obsolete PowerShell modules

    .DESCRIPTION
    There's no simple way to uninstall obsolete PowerShell modules on a system, where "obsolete" refers to a module for which a newer version is installed.

    This function is designed to make the process of uninstalling old PowerShell modules as simple as possible, instead of a manual and tedious process.

    .PARAMETER Name
    The names of the modules to be uninstalled.

    If not specified, all obsolete modules will be uninstalled.

    .PARAMETER IncludeDscModules
    Include obsolete DSC (Desired State Configuration) modules during uninstallation.

    This switch is provided as a safety check, as it's not uncommon to have multiple versions of DSC modules installed which are being actively used.

    .PARAMETER ProgressParentId
    The ID of the progress bar displayed by the calling (parent) command.

    This optional parameter can be used to ensure the progress bar is reliably displayed as a child of the progress bar displayed by a calling command.

    .EXAMPLE
    Uninstall-ObsoleteModule

    Uninstalls all obsolete PowerShell modules on the system, excluding obsolete DSC modules.

    .NOTES
    Obsolete modules which are a dependency of other modules will not be uninstalled, provided they are referenced in the dependent module(s) manifest.

    This function relies on functionality provided by the PowerShellGet module. At least PowerShellGet v2 is required, but PowerShellGet v3 is in many areas substantially more performant.

    At the time of writing, PowerShellGet v3 is still in pre-release and is not at feature parity with PowerShellGet v2. For this reason it's recommended you install both, and this function will use whichever is best for the task at hand.

    Running this function without Administrator privileges will result in only obsolete modules installed in the per-user scope being uninstalled. Obsolete modules installed for all users will output a warning message requesting Administrator privileges.

    .LINK
    https://github.com/ralish/PSWinGlue
#>

#Requires -Version 3.0

[CmdletBinding(SupportsShouldProcess)]
[OutputType([Void])]
Param(
    [ValidateNotNullOrEmpty()]
    [String[]]$Name,

    [Switch]$IncludeDscModules,

    [ValidateRange(-1, [Int]::MaxValue)]
    [Int]$ProgressParentId
)

$PowerShellGet = @(Get-Module -Name 'PowerShellGet' -ListAvailable -Verbose:$false)
if (!$PowerShellGet) {
    throw 'Required module not available: PowerShellGet'
}

$PsGetV3 = $false
$PsGetLatest = $PowerShellGet | Sort-Object -Property 'Version' -Descending | Select-Object -First 1
if ($PsGetLatest.Version.Major -ge 3) {
    $PsGetV3 = $true
} elseif ($PsGetLatest.Version.Major -lt 2) {
    throw 'At least PowerShellGet v2 is required but found: {0}' -f $PsGetLatest.Version
}
Write-Verbose -Message ('Using PowerShellGet v{0}' -f $PsGetLatest.Version)

# Not all platforms have DSC support as part of PowerShell itself
$DscSupported = Get-Command -Name 'Get-DscResource' -ErrorAction Ignore
if ($IncludeDscModules -and !$DscSupported) {
    throw 'Unable to enumerate DSC modules as Get-DscResource command unavailable.'
}

$WriteProgressParams = @{
    Activity = 'Uninstalling obsolete PowerShell modules'
}

if ($PSBoundParameters.ContainsKey('ProgressParentId')) {
    $WriteProgressParams['ParentId'] = $ProgressParentId
    $WriteProgressParams['Id'] = $ProgressParentId + 1
}

$GetParams = @{}
if ($Name) {
    $GetParams['Name'] = $Name
}

Write-Progress @WriteProgressParams -Status 'Enumerating installed modules' -PercentComplete 1
if ($PsGetV3) {
    $InstalledModules = Get-PSResource -Verbose:$false @GetParams
} else {
    $InstalledModules = Get-InstalledModule -Verbose:$false @GetParams
}

# Get-PSResource returns all module versions, while Get-InstalledModule only
# returns the latest version, so this is only necessary for PsGet v2.
$UniqueModules = $InstalledModules.Name | Sort-Object -Unique

# Percentage of the total progress for updating modules
$ProgressPercentUpdatesBase = 10
$ProgressPercentUpdatesSection = 90

if (!$IncludeDscModules -and $DscSupported) {
    Write-Progress @WriteProgressParams -Status 'Enumerating DSC modules for exclusion' -PercentComplete 5

    # Get-DscResource likes to output multiple progress bars but lacks the good
    # manners to clean them up. The result is a visual mess when when we've got
    # our own progress bars.
    $OriginalProgressPreference = $ProgressPreference
    Set-Variable -Name 'ProgressPreference' -Scope Global -Value Ignore -WhatIf:$false

    try {
        # Get-DscResource may output various errors, most often due to
        # duplicate resources. That's often the case with, for example, the
        # PackageManagement module being available in multiple locations.
        $DscModules = @(Get-DscResource -Module * -ErrorAction Ignore -Verbose:$false | Select-Object -ExpandProperty ModuleName -Unique)
    } finally {
        Set-Variable -Name 'ProgressPreference' -Scope Global -Value $OriginalProgressPreference -WhatIf:$false
    }
}

if (!$PsGetV3) {
    # Retrieve all installed modules (inc. all versions). We use this as an
    # optimisation to avoid calling Get-InstalledModule wherever possible.
    Write-Progress @WriteProgressParams -CurrentOperation 'Enumerating available modules' -PercentComplete 5
    $AvailableModules = Get-Module -ListAvailable -Verbose:$false @GetParams
}

# Uninstall obsolete modules compatible with PowerShellGet
for ($ModuleIdx = 0; $ModuleIdx -lt $UniqueModules.Count; $ModuleIdx++) {
    $ModuleName = $UniqueModules[$ModuleIdx]

    if (!$IncludeDscModules -and $DscSupported -and $ModuleName -in $DscModules) {
        Write-Verbose -Message ('Skipping DSC module: {0}' -f $ModuleName)
        continue
    }

    # Retrieve all versions of the module
    if ($PsGetV3) {
        $AllVersions = @($InstalledModules | Where-Object Name -EQ $ModuleName)
    } else {
        # Try to avoid additional calls to Get-InstalledModule as it's *very*
        # slow. Unfortunately "Get-Module -ListAvailable" can't be relied on
        # due to a bug in older PowerShell releases. Affected releases won't
        # list modules with certain names if they aren't already imported.
        #
        # See: https://github.com/PowerShell/PowerShell/pull/8777
        $MatchingModules = @($AvailableModules | Where-Object Name -EQ $ModuleName)
        if ($MatchingModules -and $MatchingModules.Count -eq 1) {
            continue
        }

        $AllVersions = @(Get-InstalledModule -Name $ModuleName -AllVersions -Verbose:$false)
    }

    # Only a single version of the module appears to be installed
    if ($AllVersions.Count -eq 1) {
        continue
    }

    $SortedModules = @($AllVersions | Sort-Object -Property Version)
    $ObsoleteModules = @($SortedModules[0..($SortedModules.Count - 2)])
    $ObsoleteVersions = $ObsoleteModules.Version -join ', '
    $ObsoleteVersionsWithModuleName = '{0}: {1}' -f $ModuleName, ($ObsoleteVersions -join ', ')

    $PercentComplete = ($ModuleIdx + 1) / $UniqueModules.Count * $ProgressPercentUpdatesSection + $ProgressPercentUpdatesBase
    $CurrentOperation = 'Uninstalling {0} version(s): {1}' -f $ModuleName, $ObsoleteVersions
    Write-Progress @WriteProgressParams -Status $CurrentOperation -PercentComplete $PercentComplete

    if ($PSCmdlet.ShouldProcess($ObsoleteVersionsWithModuleName, 'Uninstall obsolete versions')) {
        foreach ($ObsoleteModule in $ObsoleteModules) {
            try {
                if ($PsGetV3) {
                    $ObsoleteModule | Uninstall-PSResource -ErrorAction Stop -Verbose:$false
                } else {
                    $ObsoleteModule | Uninstall-Module -ErrorAction Stop -Verbose:$false
                }
            } catch {
                switch -Regex ($PSItem.FullyQualifiedErrorId) {
                    '^AdminPrivilegesRequiredForUninstall,' {
                        Write-Warning -Message ('Unable to uninstall module without Administrator rights: {0} v{1}' -f $ObsoleteModule.Name, $ObsoleteModule.Version)
                    }

                    # Uninstall-Module prints its own warning
                    '^ModuleIsInUse,' { }

                    '^(UnableToUninstallAsOtherModulesNeedThisModule|UninstallPSResourcePackageIsaDependency),' {
                        Write-Warning -Message ('Unable to uninstall module due to presence of dependent modules: {0} v{1}' -f $ObsoleteModule.Name, $ObsoleteModule.Version)
                    }

                    Default { throw }
                }
            }
        }
    }
}

Write-Progress @WriteProgressParams -Completed
