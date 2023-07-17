<#
    .SYNOPSIS
    Lexically sorts the exported values for each registry key in a Windows Registry export

    .DESCRIPTION
    Exporting a registry key to a ".reg" file will typically not result in the exported values being lexically sorted.

    This command sorts the exported values under each registry key in a Windows Registry export in lexicographical order.

    .PARAMETER Path
    The path to a Windows Registry export which will have the values for each exported registry key sorted lexically.

    .EXAMPLE
    Sort-Registryexport -Path Export.reg

    Sorts the registry values lexically for each exported registry key in the Export.reg file.

    .NOTES
    Windows Registry export are expected to have a first line beginning with "Windows Registry Editor".

    Registry keys are not sorted (only their values), however, Windows built-in tools export keys in lexical order.

    .LINK
    https://github.com/ralish/PSWinGlue
#>

#Requires -Version 3.0

[CmdletBinding()]
[OutputType([Void])]
Param(
    [Parameter(Mandatory)]
    [String]$Path
)

$RegSignature = 'Windows Registry Editor'

try {
    $RegFile = Get-Item -Path $Path -ErrorAction Stop
} catch {
    throw $_
}

if ($RegFile -isnot [IO.FileInfo]) {
    throw 'Expected a file but received: {0}' -f $RegFile.GetType().Name
}

$RegFileSig = Get-Content -Path $RegFile -TotalCount 1
if ($RegFileSig -notmatch "^$RegSignature") {
    throw 'File does not begin with expected "{0}" signature: {1}' -f $RegSignature, $RegFile.Name
}

# Everything looks good so retrieve the complete file content
$RegFileContent = Get-Content -Path $RegFile

# List to hold the new file content starting with the signature
$RegNewContent = New-Object -TypeName 'Collections.Generic.List[String]'
$RegNewContent.Add($RegFileContent[0])

# List holding entries for the current registry key (INI section)
$RegKeyContent = New-Object -TypeName 'Collections.Generic.List[String]'

for ($Idx = 1; $Idx -lt $RegFileContent.Count; $Idx++) {
    $Line = $RegFileContent[$Idx]

    # Blank line (ignored)
    if ([String]::IsNullOrWhiteSpace($Line)) {
        continue
    }

    # Registry key
    if ($Line -match '^\[') {
        # Sort and append entries from the previous registry key
        if ($RegKeyContent.Count -ne 0) {
            $RegKeyContent.Sort()
            foreach ($Entry in $RegKeyContent) {
                $RegNewContent.Add($Entry)
            }

            $RegKeyContent.Clear()
        }

        $RegNewContent.Add([String]::Empty)
        $RegNewContent.Add($Line)
        continue
    }

    # Registry value
    if ($Line -match '^[@"]') {
        # Handle values where the data is split over multiple lines
        if ($Line -match '\\$') {
            do {
                $Idx++
                $ExtraLine = $RegFileContent[$Idx]
                $Line += '{0}{1}' -f [Environment]::NewLine, $ExtraLine
            } while ($ExtraLine -match '\\$')
        }

        $RegKeyContent.Add($Line)
        continue
    }

    throw 'Unexpected content on line {0} sorting registry file: {1}' -f ($Idx + 1), $RegFile.Name
}

# Add any values from the final registry key
if ($RegKeyContent.Count -ne 0) {
    $RegKeyContent.Sort()
    foreach ($Entry in $RegKeyContent) {
        $RegNewContent.Add($Entry)
    }
}

Set-Content -Path $RegFile -Value $RegNewContent
