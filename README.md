PSWinGlue
=========

A PowerShell module consisting of an assortment of useful scripts and functions.

Requirements
------------

- PowerShell 3.0 (or later)

Installing
----------

### PowerShellGet (included with PowerShell 5.0)

The latest release of the module is published to the [PowerShell Gallery](https://www.powershellgallery.com/) for installation via the [PowerShellGet module](https://www.powershellgallery.com/GettingStarted):

```posh
Install-Module -Name PSWinGlue
```

You can find the module listing [here](https://www.powershellgallery.com/packages/PSWinGlue).

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

License
-------

All content is licensed under the terms of [The MIT License](LICENSE).
