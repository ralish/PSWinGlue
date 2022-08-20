<#
    .SYNOPSIS
    Removes orphan dependency packages

    .DESCRIPTION
    Windows maintains a local package cache of installed software to simplify operations which require access to the original installer. The package cache is typically located at "C:\ProgramData\Package Cache".

    Not all installers use the package cache, but Windows Installer (MSI files) typically do, as without a cached copy of the installer it's not possible to modify or even remove the existing installation.

    Unfortunately, the package cache may in some cases maintain installers for removed applications, growing in size as old installers are accumulated. This is especially the case for dependency packages.

    Visual Studio is an particularly prominent offender, as it relies on many MSIs which are frequently updated, but not always cleanly removed. The various .NET Core packages are the most common examples.

    This function can uninstall "orphaned" dependency packages and remove cached installers, as identified by the Find-OrphanDependencyPackages function. You use this function entirely at your own risk!

    .EXAMPLE
    Remove-OrphanDependencyPackages

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
    [Parameter(Mandatory)]
    [PSCustomObject[]]$Packages
)

$PowerShellCore = New-Object -TypeName Version -ArgumentList 6, 0
if ($PSVersionTable.PSVersion -ge $PowerShellCore -and $PSVersionTable.Platform -ne 'Win32NT') {
    throw '{0} is only compatible with Windows.' -f $MyInvocation.MyCommand.Name
}

Function Test-IsAdministrator {
    [CmdletBinding()]
    [OutputType([Boolean])]
    Param()

    $User = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
    if ($User.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        return $true
    }

    return $false
}

Function Test-PackageObjectProperties {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '')]
    [CmdletBinding()]
    [OutputType([Boolean])]
    Param(
        [Parameter(Mandatory)]
        [PSCustomObject]$Package
    )

    $ExpectedProperties = 'Name', 'Status', 'RegKey', 'CacheStatus', 'CacheFiles'
    foreach ($Property in $ExpectedProperties) {
        if ($Package.PSObject.Properties.Name -notcontains $Property) {
            throw 'Encounted a package with a missing required property: {0}' -f $Property
        }
    }

    if ($Package.Status -ne 'Orphaned') {
        return $false
    }

    if ($Package.CacheFiles.Count -gt 1) {
        Write-Warning -Message ('[{0}] Skipping package with multiple files.' -f $Package.Name)
        return $false
    }

    # Remaining checks don't apply
    if ($Package.CacheFiles.Count -eq 0) {
        return $true
    }

    if (!$Package.CacheFiles[0].Name.EndsWith('.msi')) {
        Write-Error -Message ('[{0}] Skipping package without MSI installer.' -f $Package.Name)
        return $false
    }

    return $true
}

if (!(Test-IsAdministrator) -and !$WhatIfPreference) {
    throw '{0} requires Administrator privileges.' -f $MyInvocation.MyCommand.Name
}

$MaxPackageNameSize = 0
foreach ($Package in $Packages) {
    if ($Package.Name.Length -gt $MaxPackageNameSize) {
        $MaxPackageNameSize = $Package.Name.Length
    }
}
# Negative for left justification and -2 to account for added brackets
$MaxPackageNameSize = ($MaxPackageNameSize / -1) - 2

foreach ($Package in $Packages) {
    if (!(Test-PackageObjectProperties -Package $Package)) {
        continue
    }

    $PackageNameBracketed = '[{0}]' -f $Package.Name
    $PackageOutputPrefix = "{0,$MaxPackageNameSize}" -f $PackageNameBracketed
    Write-Verbose -Message ('{0} Processing package ...' -f $PackageOutputPrefix)

    if ($Package.CacheFiles.Count -eq 1) {
        $MsiexecParams = @(
            '/u'
            ('"{0}"' -f $Package.CacheFiles[0].FullName)
            '/passive'
        )

        Start-Process -FilePath 'msiexec.exe' -ArgumentList $MsiexecParams -Wait
        if ($LASTEXITCODE -ne 0) {
            Write-Warning -Message ('{0} Uninstall returned exit code: {1}' -f $PackageOutputPrefix, $LASTEXITCODE)
            continue
        }

        Remove-Item -LiteralPath $Package.CacheFiles[0].Directory -Recurse
    }

    # Calling Remove-Item with -WhatIf on a registry key which will require
    # Administrator privileges to remove won't bypass the permissions check.
    # Simulate what would be returned if we did have required privileges so an
    # unhelpful "Requested registry access is not allowed." error is not shown.
    if ($WhatIfPreference) {
        $Key = Get-Item -LiteralPath $Package.RegKey
        Write-Host -Object ('What if: Performing the operation "Remove Key" on target "Item: {0}".' -f $Key.Name)
    } else {
        Remove-Item -LiteralPath $Package.RegKey -Recurse
    }
}
