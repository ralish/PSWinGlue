<#
    .SYNOPSIS
    Removes orphan dependency packages in the system package cache

    .DESCRIPTION
    Windows maintains a local package cache of installed software to simplify operations which require access to the original installer. The package cache is typically located at "C:\ProgramData\Package Cache".

    Not all installers use the package cache, but Windows Installer (MSI files) typically do, as without a cached copy of the installer it's not possible to modify or even remove the existing installation.

    Unfortunately, the package cache may in some cases maintain installers for removed applications, growing in size as old installers are accumulated. This is especially the case for dependency packages.

    Visual Studio is an particularly prominent offender, as it relies on many MSIs which are frequently updated, but not always cleanly removed. The various .NET Core packages are the most common examples.

    This function can uninstall "orphaned" dependency packages and remove cached installers, as identified by the Find-OrphanDependencyPackages function. You use this function entirely at your own risk!

    .PARAMETER Packages
    The orphaned packages to be removed as emitted by Find-OrphanDependencyPackages.

    Packages are filtered to only remove those meeting all of the following criteria:
    - Status is "Orphaned"
    - Has a single cached installer file or none
    - If a single cached installer file is present it is an MSI

    .EXAMPLE
    Find-OrphanDependencyPackages | Remove-OrphanDependencyPackages

    Removes orphan dependency packages found by Find-OrphanDependencyPackages.

    .NOTES
    Removal of an orphaned dependency package consists of several steps:
    - If a cached MSI installer was identified, invoke "msiexec" requesting it be uninstalled.
    - Remove the folder in the package cache which contains the installer.
    - Remove the registry data for the orphaned package.

    This function supports the "-WhatIf" parameter and you are *strongly* encouraged to use it to first verify what it's going to do. This can be done without Administrator rights for additional protection.

    .LINK
    https://github.com/ralish/PSWinGlue
#>

#Requires -Version 3.0

[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '')]
[CmdletBinding(SupportsShouldProcess)]
[OutputType([Void])]
Param(
    [Parameter(Mandatory, ValueFromPipeline)]
    [PSCustomObject[]]$Packages
)

Begin {
    $PowerShellCore = New-Object -TypeName Version -ArgumentList 6, 0
    if ($PSVersionTable.PSVersion -ge $PowerShellCore -and $PSVersionTable.Platform -ne 'Win32NT') {
        throw '{0} is only compatible with Windows.' -f $MyInvocation.MyCommand.Name
    }

    $User = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
    if (!($User.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))) {
        throw '{0} requires Administrator privileges.' -f $MyInvocation.MyCommand.Name
    }
}

Process {
    foreach ($Package in $Packages) {
        $ExpectedProperties = 'Name', 'Status', 'RegKey', 'CacheStatus', 'CacheFiles'
        foreach ($Property in $ExpectedProperties) {
            if ($Package.PSObject.Properties.Name -notcontains $Property) {
                throw 'Encounted a package with a missing required property: {0}' -f $Property
            }
        }

        if ($Package.Status -ne 'Orphaned') {
            continue
        }

        if ($Package.CacheFiles.Count -gt 1) {
            Write-Warning -Message ('[{0}] Skipping package with multiple files.' -f $Package.Name)
            continue
        }

        if ($Package.CacheFiles.Count -eq 1 -and !$Package.CacheFiles[0].Name.EndsWith('.msi')) {
            Write-Error -Message ('[{0}] Skipping package without MSI installer.' -f $Package.Name)
            continue
        }

        Write-Verbose -Message ('Processing: {0}' -f $Package.Name)

        if ($Package.CacheFiles.Count -eq 1) {
            $MsiexecParams = @(
                '/u'
                ('"{0}"' -f $Package.CacheFiles[0].FullName)
                '/passive'
            )

            Start-Process -FilePath 'msiexec.exe' -ArgumentList $MsiexecParams -Wait
            if ($LASTEXITCODE -ne 0) {
                Write-Warning -Message ('Uninstall of {0} returned exit code: {1}' -f $Package.Name, $LASTEXITCODE)
                continue
            }

            Remove-Item -LiteralPath $Package.CacheFiles[0].Directory -Recurse
        }

        # Calling Remove-Item with -WhatIf on a registry key which will require
        # Administrator privileges to remove won't bypass the permission check.
        # Simulate what would be returned if we had required privileges so a
        # "Requested registry access is not allowed." error is not shown.
        if ($WhatIfPreference) {
            $Key = Get-Item -LiteralPath $Package.RegKey
            Write-Host -Object ('What if: Performing the operation "Remove Key" on target "Item: {0}".' -f $Key.Name)
        } else {
            Remove-Item -LiteralPath $Package.RegKey -Recurse
        }
    }
}
