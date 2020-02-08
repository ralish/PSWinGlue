<#
    .SYNOPSIS
    Uninstalls a Visual Studio Tools for Office (VSTO) add-in

    .DESCRIPTION
    Launches the uninstallation of a VSTO add-in using VSTOInstaller.

    .PARAMETER ManifestPath
    Path to the VSTO manifest for the VSTO add-in to be uninstalled.

    The path can be any of:
    - A path on the local computer
    - A path to a UNC file share
    - A path to a HTTP(S) site

    .PARAMETER Silent
    Runs the VSTO uninstallation silently.

    .EXAMPLE
    Uninstall-VSTOAddin.ps1 -ManifestPath "https://example.s3.amazonaws.com/live/MyAddin.vsto" -Silent

    Silently uninstalls the VSTO add-in described by the provided manifest.

    .NOTES
    The Visual Studio 2010 Tools for Office Runtime, which includes VSTOInstaller.exe, must be present on the system.

    For a silent uninstallation the certificate with which the manifest is signed must be present in the Trusted Publishers certificate store.

    The VSTOInstaller.exe which matches the bitness of the PowerShell runtime will be used (e.g. 64-bit VSTOInstaller.exe with 64-bit PowerShell).

    The parameters for VSTOInstaller.exe are documented on TechNet:
    https://docs.microsoft.com/en-us/previous-versions/visualstudio/visual-studio-2010/bb772078(v=vs.100)

    The error codes for VSTOInstaller.exe are documented on MSDN:
    https://docs.microsoft.com/en-us/previous-versions/visualstudio/visual-studio-2010/bb757423(v=vs.100)

    .LINK
    https://github.com/ralish/PSWinGlue
#>

#Requires -Version 3.0

[CmdletBinding()]
Param(
    [Parameter(Mandatory)]
    [String]$ManifestPath,

    [Switch]$Silent
)

$VstoInstaller = '{0}\Microsoft Shared\VSTO\10.0\VSTOInstaller.exe' -f $env:CommonProgramFiles
$VstoArguments = @('/Uninstall', ('"{0}"' -f $ManifestPath))

if ($Silent) {
    $VstoArguments += '/Silent'
}

if (!(Test-Path -Path $VstoInstaller)) {
    throw 'VSTO Installer not present at: {0}' -f $VstoInstaller
}

Write-Verbose -Message ('Uninstalling VSTO add-in with arguments: {0}' -f [String]::Join(' ', $VstoArguments))
$VstoProcess = Start-Process -FilePath $VstoInstaller -ArgumentList $VstoArguments -Wait -PassThru
if ($VstoProcess.ExitCode -ne 0) {
    throw 'VSTO Installer failed with code: {0}' -f $VstoProcess.ExitCode
}
