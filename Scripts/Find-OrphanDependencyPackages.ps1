<#
    .SYNOPSIS
    Locates orphan dependency packages in the system package cache

    .DESCRIPTION
    Windows maintains a local package cache of installed software to simplify operations which require access to the original installer. The package cache is typically located at "C:\ProgramData\Package Cache".

    Not all installers use the package cache, but Windows Installer (MSI files) typically do, as without a cached copy of the installer it's not possible to modify or even remove the existing installation.

    Unfortunately, the package cache may in some cases maintain installers for removed applications, growing in size as old installers are accumulated. This is especially the case for dependency packages.

    Visual Studio is an particularly prominent offender, as it relies on many MSIs which are frequently updated, but not always cleanly removed. The various .NET Core packages are the most common examples.

    This function attempts to identify "orphaned" dependency packages, which after inspection can be removed using the Remove-OrphanDependencyPackages function. You use this function entirely at your own risk!

    .EXAMPLE
    Find-OrphanDependencyPackages

    Analyzes the registry and filesystem for orphan dependency packages.

    .NOTES
    There's no simple way to "clean" the package cache and associated registry data. The best that can be done is to try and determine if a package is unused and match registry data to a cached installer.

    The general process is to inspect the registry data for each package, and given an absence of any metadata and a "Dependents" key with no sub-keys, it's fairly safely assume it is an orphaned dependency.

    Matching a package to a cached installer is non-trivial, as there's no general way to make the match. This function is able to do so for various .NET Core packages due to the predictable naming scheme.

    .LINK
    https://github.com/ralish/PSWinGlue
#>

#Requires -Version 3.0

[CmdletBinding()]
[OutputType([PSCustomObject[]])]
Param()

$PowerShellCore = New-Object -TypeName Version -ArgumentList 6, 0
if ($PSVersionTable.PSVersion -ge $PowerShellCore -and $PSVersionTable.Platform -ne 'Win32NT') {
    throw '{0} is only compatible with Windows.' -f $MyInvocation.MyCommand.Name
}

# Package cache path
$Script:PackageCachePath = Join-Path -Path $env:ProgramData -ChildPath 'Package Cache'

# Registry MSI metadata
$Script:InstallerRegPath = 'HKLM:\SOFTWARE\Classes\Installer'
# Registry MSI dependencies
$Script:DependenciesRegPath = Join-Path -Path $InstallerRegPath -ChildPath 'Dependencies'

# Known packages
$Script:KnownPackages = @{
    dotnet_apphost_pack                                = @{
        Registry  = 'dotnet_apphost_pack_(\d+\.\d+\.\d+)_([a-z0-9_]+)'
        Directory = 'v$1'
        File      = '^dotnet-apphost-pack-.+-$2\.msi'
    }
    # The registry key name is insufficient to match to MSI files in the
    # package cache. Instead, we have to find MSI files matching the below
    # pattern, then extract a record from them which we can use to match
    # against the correct registry key.
    #Dotnet_CLI                                        = @{
    #   Registry  = 'Dotnet_CLI_(\d+\.\d+\.\d+)\.\d+_([a-z0-9]+)'
    #   Directory = 'v\d+\.\d+\.\d+'
    #   File      = '^dotnet-sdk-internal-.+-$2\.msi'
    #}
    Dotnet_CLI_HostFxr                                 = @{
        Registry  = 'Dotnet_CLI_HostFxr_(\d+\.\d+\.\d+)_([a-z0-9]+)'
        Directory = 'v$1'
        File      = '^dotnet-hostfxr-.+-$2\.msi'
    }
    Dotnet_CLI_SharedHost                              = @{
        Registry  = 'Dotnet_CLI_SharedHost_(\d+\.\d+\.\d+)_([a-z0-9]+)'
        Directory = 'v$1'
        File      = '^dotnet-host-.+-$2\.msi'
    }
    dotnet_runtime                                     = @{
        Registry  = 'dotnet_runtime_(\d+\.\d+\.\d+)_([a-z0-9]+)'
        Directory = 'v$1'
        File      = '^dotnet-runtime-.+-$2\.msi'
    }
    dotnet_targeting_pack                              = @{
        Registry  = 'dotnet_targeting_pack_(\d+\.\d+\.\d+)_([a-z0-9]+)'
        Directory = 'v$1'
        File      = '^dotnet-targeting-pack-.+-$2\.msi'
    }
    'DotNet.CLI.SharedFramework.Microsoft.NETCore.App' = @{
        Registry  = 'DotNet\.CLI\.SharedFramework\.Microsoft\.NETCore\.App_(\d+\.\d+\.\d+)_([a-z0-9]+)'
        Directory = 'v\d+\.\d+\.\d+'
        File      = '^dotnet-runtime-$1-.+-$2\.msi'
    }
    'Microsoft.AspNetCore.SharedFramework'             = @{
        Registry  = 'Microsoft\.AspNetCore\.SharedFramework_([a-z0-9]+)_.+,v(\d+\.\d+\.\d+)'
        Directory = 'v$2'
        File      = '^AspNetCoreSharedFramework-$1\.msi'
    }
    'Microsoft.AspNetCore.TargetingPack'               = @{
        Registry  = 'Microsoft\.AspNetCore\.TargetingPack_([a-z0-9]+)_.+,v(\d+\.\d+\.\d+)'
        Directory = 'v$2'
        File      = '^aspnetcore-targeting-pack-$2-.+-$1\.msi'
    }
    NetCore_Templates                                  = @{
        Registry  = 'NetCore_Templates_\d+\.\d+_(\d+\.\d+\.\d+).*_([a-z0-9]+)'
        Directory = 'v$1'
        File      = '^dotnet-\d+templates-.+-$2\.msi'
    }
    windowsdesktop_runtime                             = @{
        Registry  = 'windowsdesktop_runtime_(\d+\.\d+\.\d+)_([a-z0-9]+)'
        Directory = 'v$1'
        File      = '^windowsdesktop-runtime-.+-$2\.msi'
    }
    windowsdesktop_targeting_pack                      = @{
        Registry  = 'windowsdesktop_targeting_pack_(\d+\.\d+\.\d+)_([a-z0-9]+)'
        Directory = 'v$1'
        File      = '^windowsdesktop-targeting-pack-.+-$2\.msi'
    }
}

Function Find-DotNetCliPackagesFromCache {
    [CmdletBinding()]
    [OutputType([PSCustomObject[]])]
    Param()

    # Retrieve all Dotnet_CLI packages
    $DncFileRegex = '^dotnet-sdk-internal-.+\.msi'
    $DncInstallers = Get-ChildItem -Path $Script:PackageCachePath -File -Recurse | Where-Object Name -Match $DncFileRegex

    # Create a Windows Installer COM object
    $Msi = New-Object -ComObject 'WindowsInstaller.Installer'

    # Windows Installer method parameters
    $MsiOpenDatabaseModeReadOnly = 0
    $MsiOpenViewQuery = @('SELECT `ProviderKey` FROM `WixDependencyProvider`')

    $Results = New-Object -TypeName 'Collections.Generic.List[PSCustomObject]'
    foreach ($DncInstaller in $DncInstallers) {
        $DotNetCli = [PSCustomObject]@{
            Name     = $DncInstaller.Name
            File     = $DncInstaller
            Provider = [String]::Empty
        }

        $Results.Add($DotNetCli)

        try {
            # Open the MSI database
            $MsiOpenDatabaseParams = $DncInstaller.FullName, $MsiOpenDatabaseModeReadOnly
            $MsiDatabase = $Msi.GetType().InvokeMember('OpenDatabase', 'InvokeMethod', $null, $Msi, $MsiOpenDatabaseParams)

            # Retrieve all records from the WixDependencyProvider table. Only a
            # subset of SQL is supported and the "LIKE" clause is unfortunately
            # not included.
            $MsiView = $Msi.GetType().InvokeMember('OpenView', 'InvokeMethod', $null, $MsiDatabase, $MsiOpenViewQuery)
            $null = $MsiView.GetType().InvokeMember('Execute', 'InvokeMethod', $null, $MsiView, $null)

            # Iterate over the returned records (there should only be one)
            while ($MsiRecord = $MsiView.GetType().InvokeMember('Fetch', 'InvokeMethod', $null, $MsiView, $null)) {
                $MsiRecordValue = $MsiRecord.GetType().InvokeMember('StringData', 'GetProperty', $null, $MsiRecord, 1)
                if ($MsiRecordValue -match '^Dotnet_CLI_') {
                    $DotNetCli.Provider = $MsiRecordValue
                    break
                }
            }
        } finally {
            $null = [Runtime.InteropServices.Marshal]::ReleaseComObject($MsiRecord)
            $null = [Runtime.InteropServices.Marshal]::ReleaseComObject($MsiView)
            $null = [Runtime.InteropServices.Marshal]::ReleaseComObject($MsiDatabase)
        }

        if (!$DotNetCli.Provider) {
            Write-Warning -Message ('[{0}] Failed to find Provider value.' -f $DncInstaller.Name)
        }
    }

    return ($Results.ToArray() | Sort-Object -Property Name)
}

Function Find-OrphanDependenciesFromRegistry {
    [CmdletBinding()]
    [OutputType([PSCustomObject[]])]
    Param()

    # Retrieve all package dependencies
    $Dependencies = Get-ChildItem -Path $Script:DependenciesRegPath

    # Filter out packages with dependencies
    $Results = New-Object -TypeName 'Collections.Generic.List[PSCustomObject]'
    foreach ($DependencyKey in $Dependencies) {
        $Dependency = [PSCustomObject]@{
            Name        = $DependencyKey.PSChildName
            Status      = 'Active'
            RegKey      = $DependencyKey -replace '^HKEY_LOCAL_MACHINE', 'HKLM:'
            CacheStatus = $null
            CacheFiles  = (New-Object -TypeName 'Collections.Generic.List[IO.FileSystemInfo]')
        }

        $Results.Add($Dependency)

        # Any values (inc. default value)
        $BaseValues = Get-ItemProperty -Path $DependencyKey.PSPath
        if ($BaseValues) {
            continue
        }

        # No sub-keys
        $BaseKeys = @(Get-ChildItem -Path $DependencyKey.PSPath)
        if (!$BaseKeys) {
            continue
        }

        # Any sub-keys except "Dependents"
        if ($BaseKeys.Count -gt 1 -or $BaseKeys.PSChildName -ne 'Dependents') {
            continue
        }

        $DependentsKey = $BaseKeys[0]

        # Any values under "Dependents" (inc. default value)
        $DependentsValues = Get-ItemProperty -Path $DependentsKey.PSPath
        if ($DependentsValues) {
            continue
        }

        # Any sub-keys under "Dependents"
        $DependentsKeys = Get-ChildItem -Path $DependentsKey.PSPath
        if ($DependentsKeys) {
            continue
        }

        $Dependency.Status = 'Orphaned'
    }

    return ($Results.ToArray() | Sort-Object -Property Name)
}

Function Resolve-DotNetCliPackagesCacheToRegistry {
    [CmdletBinding()]
    [OutputType([Void])]
    Param(
        [Parameter(Mandatory)]
        [PSCustomObject[]]$DotNetCliPackages,

        [Parameter(Mandatory)]
        [PSCustomObject[]]$RegistryPackages
    )

    foreach ($DncPackage in $DotNetCliPackages) {
        $RegistryPackage = $RegistryPackages | Where-Object Name -EQ $DncPackage.Provider

        if (!$RegistryPackage) {
            Write-Warning -Message ('Unable to associate {0} to a registry package.' -f $DncPackage.Name)
            continue
        }

        if ($RegistryPackage.CacheFiles.Count -ne 0) {
            Write-Error -Message ('Registry package "{0}" already associated with Dotnet_CLI package but matched: {1}' -f $RegistryPackage.Name, $DncPackage.Name)
            continue
        }

        $RegistryPackage.CacheFiles.Add($DncPackage.File)
    }
}

Function Resolve-OrphanDependenciesRegistryToCache {
    [CmdletBinding()]
    [OutputType([Void])]
    Param(
        [Parameter(Mandatory)]
        [PSCustomObject[]]$RegistryPackages
    )

    # Retrieve all cached packages
    $Packages = Get-ChildItem -Path $Script:PackageCachePath -File -Recurse

    foreach ($RegistryPackage in $RegistryPackages) {
        $RegistryPackage.CacheStatus = 'None found'

        # Locate the "known package" entry
        $KnownPackage = $null
        if ($RegistryPackage.Name -match '^([a-z_.]+)[_.]') {
            $KnownPackageName = $Matches[1]
            if ($Script:KnownPackages.Keys -contains $KnownPackageName) {
                $KnownPackage = $Script:KnownPackages[$KnownPackageName]
            }
        }

        if (!$KnownPackage) {
            $RegistryPackage.CacheStatus = 'Not searched'
            continue
        }

        # Retrieve the match on the package registry key
        if ($RegistryPackage.Name -match $KnownPackage.Registry) {
            $PackageRegVersion = $Matches[1]
            if ($Matches.Count -ge 3) {
                $PackageRegArch = $Matches[2]
            } else {
                $PackageRegArch = [String]::Empty
            }
        } else {
            $RegistryPackage.CacheStatus = 'Error'
            Write-Warning -Message ('[{0}] Known package registry key did not match.' -f $RegistryPackage.Name)
            continue
        }

        # Filter package cache directories on the regex
        $PackageDirs = New-Object -TypeName 'Collections.Generic.List[IO.FileSystemInfo]'
        $PackageDirRegex = $KnownPackage.Directory -replace '\$1', $PackageRegVersion -replace '\$2', $PackageRegArch
        foreach ($Package in $Packages) {
            if ($Package.Directory -match $PackageDirRegex) {
                $PackageDirs.Add($Package)
            }
        }

        if ($PackageDirs.Count -eq 0) {
            continue
        }

        # Filter package cache files on the regex
        $PackageFileRegex = $KnownPackage.File -replace '\$1', $PackageRegVersion -replace '\$2', $PackageRegArch
        foreach ($PackageFile in $PackageDirs) {
            if ($PackageFile.Name -match $PackageFileRegex) {
                $RegistryPackage.CacheFiles.Add($PackageFile)
            }
        }

        if ($RegistryPackage.CacheFiles.Count -gt 0) {
            $RegistryPackage.CacheStatus = '{0} files' -f $RegistryPackage.CacheFiles.Count

            if ($RegistryPackage.CacheFiles.Count -gt 1) {
                Write-Warning -Message ('[{0}] Found {1} files in package cache.' -f $RegistryPackage.Name, $RegistryPackage.CacheFiles.Count)
            }
        }
    }
}

$Packages = Find-OrphanDependenciesFromRegistry
Resolve-OrphanDependenciesRegistryToCache -RegistryPackages $Packages

$DotNetCli = Find-DotNetCliPackagesFromCache
Resolve-DotNetCliPackagesCacheToRegistry -DotNetCliPackages $DotNetCli -RegistryPackages $Packages

return $Packages
