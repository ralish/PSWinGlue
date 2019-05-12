Changelog
=========

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