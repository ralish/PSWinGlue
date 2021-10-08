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
Write-Debug -Message ('Local Office add-ins path: {0}' -f $LocalAddinsPath)

$AddinsPath = Get-Item -Path $Path -ErrorAction Stop
if ($AddinsPath -is [IO.FileInfo]) {
    $Items = @($AddinsPath)
} elseif ($AddinsPath -is [IO.DirectoryInfo]) {
    $Items = @(Get-ChildItem -Path $AddinsPath)
} else {
    throw 'Path must be a directory or Excel add-in file.'
}

$Addins = @($Items | Where-Object Extension -In '.xla', '.xlam')
if ($Addins.Count -eq 0) {
    throw 'Provided path has no Excel add-ins (xla/xlam).'
}

try {
    Write-Debug -Message 'Creating Excel COM object ...'
    $Excel = New-Object -ComObject Excel.Application
} catch {
    throw 'Failed to create Excel COM object.'
}

try {
    Write-Debug -Message 'Retrieving Excel add-ins ...'
    $ExcelAddinsObj = $Excel.AddIns
    $ExcelAddins = New-Object -TypeName Collections.ArrayList
    foreach ($ExcelAddin in $ExcelAddinsObj) {
        $null = $ExcelAddins.Add($ExcelAddin)
    }

    foreach ($Addin in $Addins) {
        if ($ExcelAddins.Name -contains $Addin.Name) {
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

        # AddIns.Add() requires an open workbook
        if (!$ExcelWorkbook) {
            Write-Debug -Message 'Creating Excel workbook ...'
            $ExcelWorkbook = $Excel.Workbooks.Add()
        }

        Write-Debug -Message ('Adding Excel add-in: {0}' -f $Addin.Name)
        $ExcelAddin = $ExcelAddinsObj.Add($Addin.FullName, $false)
        $ExcelAddin.Installed = $true
        $null = [Runtime.InteropServices.Marshal]::FinalReleaseComObject($ExcelAddin)

        Write-Verbose -Message ('Installed Excel add-in: {0}' -f $Addin.Name)
    }
} finally {
    if ($ExcelWorkbook) {
        Write-Debug -Message 'Closing Excel workbook ...'
        $ExcelWorkbook.Close($false)
        $null = [Runtime.InteropServices.Marshal]::FinalReleaseComObject($ExcelWorkbook)
    }

    Write-Debug -Message 'Disposing Excel add-ins ...'
    foreach ($ExcelAddin in $ExcelAddins) {
        $null = [Runtime.InteropServices.Marshal]::FinalReleaseComObject($ExcelAddin)
    }
    $null = [Runtime.InteropServices.Marshal]::FinalReleaseComObject($ExcelAddinsObj)

    Write-Debug -Message 'Quitting Excel ...'
    $Excel.Quit()
    $null = [Runtime.InteropServices.Marshal]::FinalReleaseComObject($Excel)
}
