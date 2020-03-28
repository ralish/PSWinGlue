<#
    .SYNOPSIS
    Manages a Visual Studio Tools for Office (VSTO) add-in

    .DESCRIPTION
    Launches the installation or uninstallation of a VSTO add-in.

    .PARAMETER Operation
    The operation to perform on the VSTO add-in.

    Valid operations are:
    - Install
    - Uninstall

    .PARAMETER ManifestPath
    Path to the manifest file for the VSTO add-in.

    The path can be any of:
    - A path on the local computer
    - A path to a UNC file share
    - A path to a HTTP(S) site

    .PARAMETER Silent
    Runs the VSTO operation silently.

    .EXAMPLE
    Manage-VSTOAddin -Operation Install -ManifestPath "https://example.s3.amazonaws.com/live/MyAddin.vsto" -Silent

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
Param(
    [Parameter(Mandatory)]
    [ValidateSet('Install', 'Uninstall')]
    [String]$Operation,

    [Parameter(Mandatory)]
    [String]$ManifestPath,

    [Switch]$Silent
)

$VstoInstaller = '{0}\Microsoft Shared\VSTO\10.0\VSTOInstaller.exe' -f $env:CommonProgramFiles
$VstoArguments = @(('/{0}' -f $Operation), ('"{0}"' -f $ManifestPath))

if ($Silent) {
    $VstoArguments += '/Silent'
}

if (!(Test-Path -Path $VstoInstaller)) {
    throw 'VSTO Installer not present at: {0}' -f $VstoInstaller
}

Write-Verbose -Message ('{0}ing VSTO add-in with arguments: {1}' -f $Operation, [String]::Join(' ', $VstoArguments))
$VstoProcess = Start-Process -FilePath $VstoInstaller -ArgumentList $VstoArguments -Wait -PassThru
if ($VstoProcess.ExitCode -ne 0) {
    throw 'VSTO Installer failed with code: {0}' -f $VstoProcess.ExitCode
}
