[CmdletBinding()]
Param(
    [Parameter(Position=1,Mandatory=$true)]
        [String]$AddinPath,
        [Switch]$Reinstall
)

$ErrorActionPreference = 'Stop'
$ExcelAddinsPath = Join-Path $env:APPDATA 'Microsoft\AddIns'

if (Test-Path -Path $AddinPath -PathType Leaf) {
    $Addin = Get-ChildItem -Path $AddinPath
    if ($Addin.Extension -NotIn ('.xla', '.xlam')) {
        Write-Error 'File does not appear to be an Excel add-in!'
    }
} else {
    Write-Error 'Invalid add-in path provided!'
}

try {
    $Excel = New-Object -ComObject Excel.Application
} catch {
    Write-Error 'Microsoft Excel does not appear to be installed!'
}

try {
    $ExcelAddins = $Excel.Addins
    # The Add() method of the AddIns interface will fail if we don't have a workbook!
    $ExcelWorkbook = $Excel.Workbooks.Add()
    $AddinInstalled = $ExcelAddins | ? {$_.Name -eq $Addin.Name}

    if (!$AddinInstalled -or $Reinstall) {
        if (!(Test-Path -Path $ExcelAddinsPath -PathType Container)) {
            New-Item -Path $ExcelAddinsPath -ItemType Directory
        }
        Copy-Item -Path $Addin.FullName -Destination $ExcelAddinsPath -Force
        $Addin = Get-ChildItem -Path (Join-Path $ExcelAddinsPath $Addin.Name)
        $Addin.IsReadOnly = $true

        $NewAddin = $ExcelAddins.Add($Addin.FullName, $false)
        $NewAddin.Installed = $true
        Write-Host ('Add-in "' + $Addin.BaseName + '" successfully installed!')
    } else {
        Write-Host ('Add-in "' + $Addin.BaseName + '" already installed!')
    }
} finally {
    $Excel.Quit()
}
