[CmdletBinding()]
Param(
    [Parameter(Mandatory=$false)]
        [Switch]$Reinstall
)

# Ensure that any errors we receive are considered fatal
$ErrorActionPreference = 'Stop'

# The path to the Registry key containing the IRESS installation details
$IressInstallRegPath = 'HKLM:\Software\Wow6432Node\DFS\IRESS\File'
# The name of the Registry property that gives the installation location
$IressInstallDirProp = 'InstallDir'

if (Test-Path -Path $IressInstallRegPath -PathType Container) {
    $IressExcelAddinsPath = Join-Path (Get-ItemProperty -Path $IressInstallRegPath).$IressInstallDirProp 'ExcelAddins'
    if (!(Test-Path -Path $IressExcelAddinsPath -PathType Container)) {
        Write-Error 'The IRESS installation on this system appears to be damaged.'
    }
} else {
    Write-Error 'IRESS does not appear to be installed on this system.'
}

$IressExcelAddins = Get-ChildItem -Path $IressExcelAddinsPath
foreach ($Addin in $IressExcelAddins) {
    Write-Host ('Installing Excel add-in: ' + $Addin.Name)
    & (Join-Path $PSScriptRoot 'Install-ExcelAddin') -AddinPath $Addin.FullName @PSBoundParameters
}
