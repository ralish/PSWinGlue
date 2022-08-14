<#
    .SYNOPSIS
    Install a Visual Studio Tools for Office (VSTO) add-in

    .DESCRIPTION
    Launches the installation of a VSTO add-in.

    .PARAMETER ManifestPath
    Path to the manifest file for the VSTO add-in.

    The path can be any of:
    - A path on the local computer
    - A path to a UNC file share
    - A path to a HTTP(S) site

    .PARAMETER Silent
    Runs the VSTO operation silently.

    .EXAMPLE
    Install-VSTOAddin -ManifestPath 'https://example.s3.amazonaws.com/live/MyAddin.vsto' -Silent

    Silently installs the VSTO add-in described by the provided manifest.

    .NOTES
    The Visual Studio 2010 Tools for Office Runtime, which includes VSTOInstaller.exe, must be present on the system.

    Silent mode requires the certificate with which the manifest is signed to be present in the Trusted Publishers certificate store.

    The VSTOInstaller.exe which matches the bitness of the PowerShell runtime will be used (e.g. 64-bit VSTOInstaller.exe with 64-bit PowerShell).

    Parameters for VSTOInstaller.exe
    https://docs.microsoft.com/en-us/previous-versions/visualstudio/visual-studio-2010/bb772078(v=vs.100)#parameters-for-vstoinstallerexe

    VSTOInstaller Error Codes
    https://docs.microsoft.com/en-us/previous-versions/visualstudio/visual-studio-2010/bb757423(v=vs.100)

    .LINK
    https://github.com/ralish/PSWinGlue
#>

#Requires -Version 3.0

[CmdletBinding()]
[OutputType()]
Param(
    [Parameter(Mandatory)]
    [String]$ManifestPath,

    [Switch]$Silent
)

$PowerShellCore = New-Object -TypeName Version -ArgumentList 6, 0
if ($PSVersionTable.PSVersion -ge $PowerShellCore -and $PSVersionTable.Platform -ne 'Win32NT') {
    throw '{0} is only compatible with Windows.' -f $MyInvocation.MyCommand.Name
}

$VstoInstaller = '{0}\Microsoft Shared\VSTO\10.0\VSTOInstaller.exe' -f $env:CommonProgramFiles
$VstoArguments = @('/Install', ('"{0}"' -f $ManifestPath))

if ($Silent) {
    $VstoArguments += '/Silent'
}

if (!(Test-Path -Path $VstoInstaller)) {
    throw 'VSTO Installer not present at: {0}' -f $VstoInstaller
}

Write-Verbose -Message ('Installing VSTO add-in with arguments: {0}', [String]::Join(' ', $VstoArguments))
$VstoProcess = Start-Process -FilePath $VstoInstaller -ArgumentList $VstoArguments -Wait -PassThru
if ($VstoProcess.ExitCode -ne 0) {
    throw 'VSTO Installer failed with code: {0}' -f $VstoProcess.ExitCode
}
