#Requires -Version 3.0

[CmdletBinding()]
Param(
    [Parameter(Mandatory)]
    [String]$Path,

    [Switch]$Copy,
    [Switch]$Reinstall
)

# Local add-ins path when copying
$LocalAddinsPath = Join-Path -Path $env:APPDATA -ChildPath 'Microsoft\AddIns'

$AddinsPath = Get-Item -Path $Path -ErrorAction Stop
if ($AddinsPath -is [IO.FileInfo]) {
    $Items = @($AddinsPath)
} elseif ($AddinsPath -is [IO.DirectoryInfo]) {
    $Items = @(Get-ChildItem -Path $AddinsPath)
} else {
    throw 'Path must be a directory or Excel add-in file.'
}

$Addins = @($Items | Where-Object Extension -in ('.xla', '.xlam'))
if ($Addins.Count -eq 0) {
    throw 'Provided path has no Excel add-ins (xla/xlam).'
}

try {
    Write-Debug -Message 'Creating Excel COM object ...'
    $Excel = New-Object -ComObject Excel.Application
} catch {
    throw 'Unable to instantiate Excel COM object.'
}

try {
    $ExcelAddins = $Excel.Addins
    # The Addins.Add() method fails without an open workbook
    $null = $Excel.Workbooks.Add()

    foreach ($Addin in $Addins) {
        $AddinInstalled = $ExcelAddins | Where-Object Name -eq $Addin.Name
        if ($AddinInstalled) {
            Write-Verbose -Message ('Excel add-in already installed: {0}' -f $Addin.Name)
            if (!$Reinstall) {
                continue
            }
        }

        if ($Copy) {
            if (!(Test-Path -Path $LocalAddinsPath -PathType Container)) {
                try {
                    $null = New-Item -Path $LocalAddinsPath -ItemType Directory -Force -ErrorAction Stop
                } catch {
                    throw ('Unable to create local Office add-ins directory: {0}' -f $LocalAddinsPath)
                }
            }

            try {
                Copy-Item -Path $Addin.FullName -Destination $LocalAddinsPath -Force -ErrorAction Stop
            } catch {
                throw ('Unable to copy add-in to local Office add-ins directory: {0}' -f $Addin.Name)
            }
        }

        # https://docs.microsoft.com/en-us/office/vba/api/excel.addins.add
        $ExcelAddin = $ExcelAddins.Add($Addin.FullName, $false)
        $ExcelAddin.Installed = $true

        Write-Verbose -Message ('Installed Excel add-in: {0}' -f $Addin.Name)
    }
} finally {
    Write-Debug -Message 'Disposing Excel COM object ...'
    $Excel.Quit()
}
