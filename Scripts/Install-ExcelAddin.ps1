<#
    .SYNOPSIS
    Installs Excel add-ins without user interaction

    .DESCRIPTION
    This function is designed to assist with handling two common installation scenarios for Excel add-ins:

    - A program which provides an Excel add-in is installed for all users but doesn't handle installation of the add-in for all users
    - A custom Excel add-in needs to be installed for all users (e.g. from a network share) without its own separate installer system

    In both scenarios the challenge is in silently automating the installation due to the interactive nature of the Excel object model.

    .PARAMETER Path
    The path to an individual Excel add-in to install, or a directory containing Excel add-ins to install.

    .PARAMETER Copy
    Copy each add-in to the Office add-ins path on the computer for the running user and perform installation from this path.

    This switch is primarily designed for use when add-in(s) reside on network storage and so may not always be accessible.

    If ommitted, add-ins will be installed directly from the provided path under the assumption it will remain accessible.

    .PARAMETER Reinstall
    Reinstall add-ins for which an existing add-in with the same file name and extension is already installed.

    This switch is especially useful with the Copy switch to ensure the latest version of add-in(s) are installed.

    If ommitted, add-ins with a file name and extension which match an existing installed add-in will be skipped.

    .EXAMPLE
    Install-ExcelAddin -Path 'C:\Program Files\Acme Inc\Excel Addins' -Reinstall

    Install Excel add-ins found in the "C:\Program Files\Acme Inc\Excel Addins" directory.

    .NOTES
    Excel add-ins must have one of the following file extensions:
    - XLA               Excel 97-2003 Add-in
    - XLAM              Excel Add-in

    .LINK
    https://github.com/ralish/PSWinGlue
#>

#Requires -Version 3.0

[CmdletBinding()]
[OutputType([Void])]
Param(
    [Parameter(Mandatory)]
    [String]$Path,

    [Switch]$Copy,
    [Switch]$Reinstall
)

$PowerShellCore = New-Object -TypeName Version -ArgumentList 6, 0
if ($PSVersionTable.PSVersion -ge $PowerShellCore -and $PSVersionTable.Platform -ne 'Win32NT') {
    throw '{0} is only compatible with Windows.' -f $MyInvocation.MyCommand.Name
}

# Valid add-in extensions
$ValidExts = @('.xla', '.xlam')

$AddinsPath = Get-Item -Path $Path -ErrorAction Stop
if ($AddinsPath -is [IO.FileInfo]) {
    if ($AddinsPath.Extension -notin $ValidExts) {
        throw 'File is not an Excel add-in (xla/xlam).'
    }

    $Addins = @($AddinsPath)
} elseif ($AddinsPath -is [IO.DirectoryInfo]) {
    $Addins = @(Get-ChildItem -Path $AddinsPath | Where-Object Extension -In $ValidExts)

    if ($Addins.Count -eq 0) {
        Write-Warning -Message 'Directory has no Excel add-ins (xla/xlam).'
        return
    }
} else {
    throw 'Path must be an add-in file or a directory containing add-ins.'
}

$OfficeAddinsPath = Join-Path -Path $env:APPDATA -ChildPath 'Microsoft\AddIns'
Write-Debug -Message ('Office add-ins path: {0}' -f $OfficeAddinsPath)

try {
    Write-Debug -Message 'Creating Excel COM object ...'
    $Excel = New-Object -ComObject 'Excel.Application'
} catch {
    throw $_
}

try {
    $ExcelAddins = $null
    $ExcelWorkbooks = $null
    $ExcelWorkbook = $null

    # Stores enumerated Excel add-ins for improved performance
    $ExcelAddinsList = New-Object -TypeName 'Collections.Generic.List[Object]'

    Write-Debug -Message 'Retrieving Excel add-ins ...'
    $ExcelAddins = $Excel.AddIns
    # The add-ins list exposed by the Excel object model is indexed from one!
    for ($i = 1; $i -le $ExcelAddins.Count; $i++) {
        $null = $ExcelAddinsList.Add($ExcelAddins.Item($i))
    }

    foreach ($Addin in $Addins) {
        if ($ExcelAddinsList.Name -contains $Addin.Name) {
            Write-Verbose -Message ('Excel add-in already installed: {0}' -f $Addin.Name)
            if (!$Reinstall) {
                continue
            }
        }

        if ($Copy) {
            if (!(Test-Path -Path $OfficeAddinsPath -PathType Container)) {
                try {
                    $null = New-Item -Path $OfficeAddinsPath -ItemType Directory -Force -ErrorAction Stop
                } catch {
                    throw ('Unable to create Office add-ins directory: {0}' -f $OfficeAddinsPath)
                }
            }

            try {
                Copy-Item -Path $Addin.FullName -Destination $OfficeAddinsPath -Force -ErrorAction Stop
            } catch {
                throw ('Unable to copy add-in to Office add-ins directory: {0}' -f $Addin.Name)
            }
        }

        # Excel.AddIns.Add() requires an open workbook. We'll only open one if
        # we find an add-in to install, which we'll re-use for any subsequent
        # add-ins to install in this run.
        if (!$ExcelWorkbook) {
            if (!$ExcelWorkbooks) {
                Write-Debug -Message 'Retrieving open workbooks ...'
                $ExcelWorkbooks = $Excel.Workbooks
            }

            Write-Debug -Message 'Creating new workbook ...'
            $ExcelWorkbook = $ExcelWorkbooks.Add()
        }

        Write-Debug -Message ('Adding Excel add-in: {0}' -f $Addin.Name)
        $ExcelAddin = $ExcelAddins.Add($Addin.FullName, $false)
        $ExcelAddin.Installed = $true
        $null = [Runtime.InteropServices.Marshal]::ReleaseComObject($ExcelAddin)

        Write-Verbose -Message ('Installed Excel add-in: {0}' -f $Addin.Name)
    }
} finally {
    if ($ExcelAddinsList.Count -gt 0) {
        foreach ($ExcelAddin in $ExcelAddinsList) {
            $null = [Runtime.InteropServices.Marshal]::ReleaseComObject($ExcelAddin)
        }
    }
    $null = [Runtime.InteropServices.Marshal]::ReleaseComObject($ExcelAddins)

    if ($ExcelWorkbook) {
        Write-Debug -Message 'Closing Excel workbook ...'
        $ExcelWorkbook.Close($false)
        $null = [Runtime.InteropServices.Marshal]::ReleaseComObject($ExcelWorkbook)
    }

    if ($ExcelWorkbooks) {
        $null = [Runtime.InteropServices.Marshal]::ReleaseComObject($ExcelWorkbooks)
    }

    Write-Debug -Message 'Quitting Excel ...'
    $Excel.Quit()
    $null = [Runtime.InteropServices.Marshal]::ReleaseComObject($Excel)
}
