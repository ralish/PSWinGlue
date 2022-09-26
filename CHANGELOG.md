Changelog
=========

v0.6.3
------

- `Uninstall-ObsoleteModule`: Retain latest `PowerShellGet` v2 & v3 versions
- `Uninstall-ObsoleteModule`: Fix `-Name` bug when module has a single version

v0.6.2
------

- `Uninstall-ObsoleteModule`: Fix incorrect `ProgressPreference` value

v0.6.1
------

- `Uninstall-ObsoleteModule`: Improve `PowerShellGet` version detection

v0.6.0
------

New functions:

- `Find-OrphanDependencyPackages`
- `Install-VSTOAddin` (previously `Manage-VSTOAddin`)
- `Remove-OrphanDependencyPackages`
- `Uninstall-VSTOAddin` (previously `Manage-VSTOAddin`)

Removed functions:

- `Manage-VSTOAddin` (split into `Install` and `Uninstall` functions)

Added built-in help:

- `Add-VpnCspConnection`
- `Get-InstalledPrograms`
- `Hide-SilverlightUpdates`
- `Install-ExcelAddin`
- `Register-MicrosoftUpdate`
- `Update-GitRepository`

Major code clean-up:

- `Add-VpnCspConnection`
- `Get-Fonts`
- `Hide-SilverlightUpdates`
- `Install-ExcelAddin`
- `Install-Font`
- `Register-MicrosoftUpdate`
- `Remove-AlternateDateStream`
- `Uninstall-Font`
- `Uninstall-ObsoleteModule`
- `Update-GitRepository`

Minor code clean-up:

- `Get-InstalledPrograms`
- `Get-TaskSchedulerEvent`
- `Set-SharedPCMode`
- `Update-OneDriveSetup`

Additional changes:

- `ConsoleAPI`: Add all missing function signatures
- `ConsoleAPI`: Add namespace (`PSWinGlue`) and class (`ConsoleAPI`)
- `Uninstall-ObsoleteModule`: Add PowerShellGet v3 support (currently in beta)
- Add check we're running on Windows for applicable functions
- Replace `-RunAsAdministrator` with equivalent check in functions
- Minor code clean-up & developer tooling improvements

v0.5.7
------

- `Uninstall-ObsoleteModule`: Replace `ValidateRangeKind` attribute for PowerShell < 6.2 compatibility

v0.5.6
------

- `Uninstall-ObsoleteModule`: Add progress bar support

v0.5.5
------

- `Get-InstalledPrograms`: Fix incorrect access to `$PSVersionTable` causing errors under PowerShell Core

v0.5.4
------

- `Get-InstalledPrograms`: Fall back to last write time of uninstall registry key

v0.5.3
------

- `Hide-SilverlightUpdates`: Sets all Silverlight updates from Windows Update to Hidden

v0.5.2
------

- `Uninstall-ObsoleteModule`: Add workaround for `Get-Module` ignoring modules with names matching locales
- `Uninstall-ObsoleteModule`: Don't throw an exception on `Uninstall-Module` returning `ModuleIsInUse`
- Apply code formatting

v0.5.1
------

- `Update-OneDriveSetup`: Quote default user profile hive path for reg.exe
- Syntax fixes for older PowerShell versions
- Performance optimisations around array use

v0.5.0
------

- Added `Add-VpnCspConnection` to add a VPN connection via the VPNv2 CSP
- Added `ConsoleAPI.ps1` to add type with P/Invoke signatures for many native Console API functions
- Added `Get-Fonts` to retrieve system or per-user fonts
- Added `Install-ExcelAddin` to install Excel add-ins
- Added `Manage-VSTOAddin` to install or uninstall VSTO add-ins
- Added `Set-SharedPCMode` to configure Windows 10 Shared PC Mode
- Added `Uninstall-Font` to uninstall system or per-user fonts
- Added `Update-OneDriveSetup` to update the in-box OneDrive setup
- `Get-ControlledGpoStatus`: Added built-in command help
- `Get-ControlledGpoStatus`: Support for custom domains & AGPM servers
- `Get-TaskSchedulerEvent`: Fix incorrect built-in help
- Remove unneeded files from published package
- Minor documentation updates & miscellaneous fixes

v0.4.6
------

- `Get-ControlledGpoStatus`: Matching of GPOs is now done by GUID instead of name
- `Get-ControlledGpoStatus`: Adds check that GPO name and any WMI filter is in sync
- `Get-ControlledGpoStatus`: Computer & user policy version mismatches are now reported separately

v0.4.5
------

- `Get-InstalledPrograms`: Check if user `Uninstall` key exists

v0.4.4
------

- Added `#Requires -Version` statement to all scripts
- `Uninstall-ObsoleteModule`: Check `PowerShellGet` module is present

v0.4.3
------

- `Install-Font`: Overhauled the script and now supports per-user fonts in Windows 10 1809 (or newer)

v0.4.2
------

- `Uninstall-ObsoleteModule`: Fixed incorrect object type under strict mode in certain scenarios

v0.4.1
------

- Added `Uninstall-ObsoleteModule` to uninstall obsolete modules installed by `PowerShellGet`

v0.4
----

- Initial stable release
