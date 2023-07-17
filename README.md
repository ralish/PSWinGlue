PSWinGlue
=========

[![pwsh ver](https://img.shields.io/powershellgallery/v/PSWinGlue)](https://www.powershellgallery.com/packages/PSWinGlue)
[![pwsh dl](https://img.shields.io/powershellgallery/dt/PSWinGlue)](https://www.powershellgallery.com/packages/PSWinGlue)
[![license](https://img.shields.io/github/license/ralish/PSWinGlue)](https://choosealicense.com/licenses/mit/)

A PowerShell module consisting of an assortment of useful scripts.

- [Requirements](#requirements)
- [Installing](#installing)
- [Usage](#usage)
- [Functions](#functions)
- [License](#license)

Requirements
------------

- PowerShell 3.0 (or later)  
  Some functions require a later PowerShell version

Installing
----------

### PowerShellGet (included with PowerShell 5.0)

The module is published to the [PowerShell Gallery](https://www.powershellgallery.com/packages/PSWinGlue):

```posh
Install-Module -Name PSWinGlue
```

### ZIP File

Download the [ZIP file](https://github.com/ralish/PSWinGlue/archive/stable.zip) of the latest release and unpack it to one of the following locations:

- Current user: `C:\Users\<your.account>\Documents\WindowsPowerShell\Modules\PSWinGlue`
- All users: `C:\Program Files\WindowsPowerShell\Modules\PSWinGlue`

### Git Clone

You can also clone the repository into one of the above locations if you'd like the ability to easily update it via Git.

### Did it work?

You can check that PowerShell is able to locate the module by running the following at a PowerShell prompt:

```posh
Get-Module PSWinGlue -ListAvailable
```

Usage
-----

This module has been written to support two methods of using the functions it includes:

- As a regular PowerShell module (i.e. install module and call the commands it exports)
- Calling individual functions directly via their script file independent of the module

To support the latter, each exported function resides in its own script file and has no dependencies on code elsewhere in the module. This allows for easy usage of the functions in scenarios where it may not be desirable to install the entire module (e.g. logon scripts or in other automated contexts). A given script can simply be copied to the desired location and directly called, typically with no additional setup beyond any documented external dependencies.

Functions
---------

### Add-VpnCspConnection

Adds a VPN connection using the VPNv2 CSP via the MDM Bridge WMI Provider.

<table>
  <tr>
    <td>Supported OS(s)</td>
    <td>Windows 10 1607 or later</td>
  </tr>
  <tr>
    <td>Minimum PowerShell version</td>
    <td>5.1</td>
  </tr>
  <tr>
    <td>Required 3rd-party module(s)</td>
    <td>None</td>
  </tr>
  <tr>
    <td>Required 3rd-party software</td>
    <td>None</td>
  </tr>
</table>

### Find-OrphanDependencyPackages

Locates orphan dependency packages in the system package cache.

<table>
  <tr>
    <td>Supported OS(s)</td>
    <td>Windows</td>
  </tr>
  <tr>
    <td>Minimum PowerShell version</td>
    <td>3.0</td>
  </tr>
  <tr>
    <td>Required 3rd-party module(s)</td>
    <td>None</td>
  </tr>
  <tr>
    <td>Required 3rd-party software</td>
    <td>None</td>
  </tr>
</table>

### Get-ControlledGpoStatus

Check Windows domain GPOs and AGPM server controlled GPOs are in sync.

<table>
  <tr>
    <td>Supported OS(s)</td>
    <td>Windows</td>
  </tr>
  <tr>
    <td>Minimum PowerShell version</td>
    <td>3.0</td>
  </tr>
  <tr>
    <td>Required 3rd-party module(s)</td>
    <td>GroupPolicy<br>Microsoft.Agpm</td>
  </tr>
  <tr>
    <td>Required 3rd-party software</td>
    <td>None</td>
  </tr>
</table>

### Get-Fonts

Retrieves registered fonts.

<table>
  <tr>
    <td>Supported OS(s)</td>
    <td>Windows</td>
  </tr>
  <tr>
    <td>Minimum PowerShell version</td>
    <td>3.0</td>
  </tr>
  <tr>
    <td>Required 3rd-party module(s)</td>
    <td>None</td>
  </tr>
  <tr>
    <td>Required 3rd-party software</td>
    <td>None</td>
  </tr>
</table>

### Get-InstalledPrograms

Retrieves installed programs.

<table>
  <tr>
    <td>Supported OS(s)</td>
    <td>Windows</td>
  </tr>
  <tr>
    <td>Minimum PowerShell version</td>
    <td>3.0</td>
  </tr>
  <tr>
    <td>Required 3rd-party module(s)</td>
    <td>None</td>
  </tr>
  <tr>
    <td>Required 3rd-party software</td>
    <td>None</td>
  </tr>
</table>

### Get-TaskSchedulerEvent

Retrieves events matching the specified IDs from the Task Scheduler event log.

<table>
  <tr>
    <td>Supported OS(s)</td>
    <td>Windows</td>
  </tr>
  <tr>
    <td>Minimum PowerShell version</td>
    <td>3.0</td>
  </tr>
  <tr>
    <td>Required 3rd-party module(s)</td>
    <td>None</td>
  </tr>
  <tr>
    <td>Required 3rd-party software</td>
    <td>None</td>
  </tr>
</table>

### Hide-SilverlightUpdates

Hides Silverlight updates.

<table>
  <tr>
    <td>Supported OS(s)</td>
    <td>Windows</td>
  </tr>
  <tr>
    <td>Minimum PowerShell version</td>
    <td>3.0</td>
  </tr>
  <tr>
    <td>Required 3rd-party module(s)</td>
    <td>None</td>
  </tr>
  <tr>
    <td>Required 3rd-party software</td>
    <td>None</td>
  </tr>
</table>

### Install-ExcelAddin

Installs Excel add-ins.

<table>
  <tr>
    <td>Supported OS(s)</td>
    <td>Windows</td>
  </tr>
  <tr>
    <td>Minimum PowerShell version</td>
    <td>3.0</td>
  </tr>
  <tr>
    <td>Required 3rd-party module(s)</td>
    <td>None</td>
  </tr>
  <tr>
    <td>Required 3rd-party software</td>
    <td>Microsoft Excel</td>
  </tr>
</table>

### Install-Font

Installs a specific font or all fonts from a directory.

<table>
  <tr>
    <td>Supported OS(s)</td>
    <td>Windows</td>
  </tr>
  <tr>
    <td>Minimum PowerShell version</td>
    <td>4.0</td>
  </tr>
  <tr>
    <td>Required 3rd-party module(s)</td>
    <td>None</td>
  </tr>
  <tr>
    <td>Required 3rd-party software</td>
    <td>None</td>
  </tr>
</table>

### Install-VSTOAddin

Install a Visual Studio Tools for Office (VSTO) add-in.

<table>
  <tr>
    <td>Supported OS(s)</td>
    <td>Windows</td>
  </tr>
  <tr>
    <td>Minimum PowerShell version</td>
    <td>3.0</td>
  </tr>
  <tr>
    <td>Required 3rd-party module(s)</td>
    <td>None</td>
  </tr>
  <tr>
    <td>Required 3rd-party software</td>
    <td>Microsoft Visual Studio Tools for Office (VSTO) Runtime</td>
  </tr>
</table>

### Register-MicrosoftUpdate

Register the Microsoft Update service with the Windows Update Agent.

<table>
  <tr>
    <td>Supported OS(s)</td>
    <td>Windows</td>
  </tr>
  <tr>
    <td>Minimum PowerShell version</td>
    <td>3.0</td>
  </tr>
  <tr>
    <td>Required 3rd-party module(s)</td>
    <td>None</td>
  </tr>
  <tr>
    <td>Required 3rd-party software</td>
    <td>None</td>
  </tr>
</table>

### Remove-AlternateDataStream

Remove common unwanted alternate data streams from files.

<table>
  <tr>
    <td>Supported OS(s)</td>
    <td>Windows</td>
  </tr>
  <tr>
    <td>Minimum PowerShell version</td>
    <td>3.0</td>
  </tr>
  <tr>
    <td>Required 3rd-party module(s)</td>
    <td>None</td>
  </tr>
  <tr>
    <td>Required 3rd-party software</td>
    <td>None</td>
  </tr>
</table>

### Remove-OrphanDependencyPackages

Removes orphan dependency packages in the system package cache.

<table>
  <tr>
    <td>Supported OS(s)</td>
    <td>Windows</td>
  </tr>
  <tr>
    <td>Minimum PowerShell version</td>
    <td>3.0</td>
  </tr>
  <tr>
    <td>Required 3rd-party module(s)</td>
    <td>None</td>
  </tr>
  <tr>
    <td>Required 3rd-party software</td>
    <td>None</td>
  </tr>
</table>

### Set-SharedPCMode

Configures Shared PC Mode using the SharedPC CSP via the MDM Bridge WMI Provider.

<table>
  <tr>
    <td>Supported OS(s)</td>
    <td>Windows 10 1607 or later</td>
  </tr>
  <tr>
    <td>Minimum PowerShell version</td>
    <td>5.1</td>
  </tr>
  <tr>
    <td>Required 3rd-party module(s)</td>
    <td>None</td>
  </tr>
  <tr>
    <td>Required 3rd-party software</td>
    <td>None</td>
  </tr>
</table>

### Sort-RegistryExport

Lexically sorts the exported values for each registry key in a Windows Registry export.

<table>
  <tr>
    <td>Supported OS(s)</td>
    <td>Windows</td>
  </tr>
  <tr>
    <td>Minimum PowerShell version</td>
    <td>3.0</td>
  </tr>
  <tr>
    <td>Required 3rd-party module(s)</td>
    <td>None</td>
  </tr>
  <tr>
    <td>Required 3rd-party software</td>
    <td>None</td>
  </tr>
</table>

### Uninstall-Font

Uninstalls a specific font by name.

<table>
  <tr>
    <td>Supported OS(s)</td>
    <td>Windows</td>
  </tr>
  <tr>
    <td>Minimum PowerShell version</td>
    <td>3.0</td>
  </tr>
  <tr>
    <td>Required 3rd-party module(s)</td>
    <td>None</td>
  </tr>
  <tr>
    <td>Required 3rd-party software</td>
    <td>None</td>
  </tr>
</table>

### Uninstall-ObsoleteModule

Uninstalls obsolete PowerShell modules.

<table>
  <tr>
    <td>Supported OS(s)</td>
    <td>Linux, macOS, Windows</td>
  </tr>
  <tr>
    <td>Minimum PowerShell version</td>
    <td>3.0</td>
  </tr>
  <tr>
    <td>Required 3rd-party module(s)</td>
    <td>PowerShellGet</td>
  </tr>
  <tr>
    <td>Required 3rd-party software</td>
    <td>None</td>
  </tr>
</table>

### Uninstall-VSTOAddin

Uninstall a Visual Studio Tools for Office (VSTO) add-in.

<table>
  <tr>
    <td>Supported OS(s)</td>
    <td>Windows</td>
  </tr>
  <tr>
    <td>Minimum PowerShell version</td>
    <td>3.0</td>
  </tr>
  <tr>
    <td>Required 3rd-party module(s)</td>
    <td>None</td>
  </tr>
  <tr>
    <td>Required 3rd-party software</td>
    <td>Microsoft Visual Studio Tools for Office (VSTO) Runtime</td>
  </tr>
</table>

### Update-GitRepository

Updates a Git repository.

<table>
  <tr>
    <td>Supported OS(s)</td>
    <td>Windows</td>
  </tr>
  <tr>
    <td>Minimum PowerShell version</td>
    <td>3.0</td>
  </tr>
  <tr>
    <td>Required 3rd-party module(s)</td>
    <td>None</td>
  </tr>
  <tr>
    <td>Required 3rd-party software</td>
    <td>Git</td>
  </tr>
</table>

### Update-OneDriveSetup

Update OneDrive during Windows image creation.

<table>
  <tr>
    <td>Supported OS(s)</td>
    <td>Windows</td>
  </tr>
  <tr>
    <td>Minimum PowerShell version</td>
    <td>3.0</td>
  </tr>
  <tr>
    <td>Required 3rd-party module(s)</td>
    <td>None</td>
  </tr>
  <tr>
    <td>Required 3rd-party software</td>
    <td>Microsoft OneDrive</td>
  </tr>
</table>

License
-------

All content is licensed under the terms of [The MIT License](LICENSE).
