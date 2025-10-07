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

try {
    $RegFile = Get-Item -Path $Path -ErrorAction Stop
} catch {
    throw $_
}

if ($RegFile -isnot [IO.FileInfo]) {
    throw 'Expected a file but received: {0}' -f $RegFile.GetType().Name
}

# Expected registry export file header
$FileHeader = 'Windows Registry Editor Version 5.00'
$RegexFileHeader = '^{0}$' -f [Regex]::Escape($FileHeader)

$RegFileHeader = Get-Content -Path $RegFile -TotalCount 1
if ($RegFileHeader -notmatch $RegexFileHeader) {
    throw 'File does not begin with expected "{0}" header: {1}' -f $FileHeader, $RegFile.Name
}

# Comment line
$RegexComment = '^\s*(;.*)\s*$'
# Registry section
#
# Registry key names can include square brackets. They're not escaped, but
# there's no other content apart from the enclosing square brackets.
$RegexSection = '^\s*(\[.+\])\s*$'
# Registry value
#
# Registry values are always quoted, except for the unnamed (default) value,
# which if set is denoted by an "@" symbol.
$RegexValue = '^\s*[@"]'
# Named registry value
#
# Registry value names can include double quotes and backslashes, both of which
# are escaped. The match group excludes the enclosing double quotes.
$RegexNamedValue = '^\s*"((?:[^"\\]|\\.)+)"'
# Multi-line registry value data
#
# Some registry value data may span multiple lines (e.g. REG_BINARY values),
# which is denoted by a single trailing backslash.
$RegexMultilineValue = '\\\s*$'

$FileContent = Get-Content -Path $Path -ErrorAction Stop
$SortedContent = New-Object -TypeName 'Collections.Generic.List[String]'

$SectionName = [String]::Empty
$SectionContent = New-Object -TypeName 'Collections.Generic.Dictionary[String, String]'

$RegComment = New-Object -TypeName 'Collections.Generic.List[String]'
$RegValue = New-Object -TypeName 'Collections.Generic.List[String]'

# Add the registry file header
$SortedContent.Add($RegFileHeader)

# Add any comments preceding the first section
for ($Idx = 1; $Idx -lt $FileContent.Count; $Idx++) {
    $Line = $FileContent[$Idx]

    if ($Line -notmatch $RegexSection) {
        $SortedContent.Add($Line)
        continue
    }

    $FirstSectionIdx = $Idx
    break
}

# Process each registry section
for ($Idx = $FirstSectionIdx; $Idx -lt $FileContent.Count; $Idx++) {
    $Line = $FileContent[$Idx]

    # Skip blank lines
    if ([String]::IsNullOrWhiteSpace($Line)) { continue }

    # Registry comment
    if ($Line -match $RegexComment) {
        if ($RegComment.Count -gt 0) {
            $RegComment += '{0}{1}' -f [Environment]::NewLine, $Matches[1]
        } else {
            $RegComment = $Matches[1]
        }

        continue
    }

    # Registry section (key path)
    if ($Line -match $RegexSection) {
        $SectionName = $Matches[1]

        # Add any comments preceding the section
        if ($RegComment) {
            $SortedContent.Add($RegComment)
            $RegComment = [String]::Empty
        }

        # Add registry values for the section
        foreach ($ValueName in ($SectionContent.Keys | Sort-Object)) {
            $SortedContent.Add($SectionContent[$ValueName])
        }
        $SectionContent.Clear()

        if ($Idx -ne $FirstSectionIdx) {
            $SortedContent.Add([String]::Empty)
        }

        # Add the start of the new section
        $SortedContent.Add($SectionName)
        continue
    }

    # Registry value (name, type, data)
    if ($Line -match $RegexValue) {
        if ($Line -match $RegexNamedValue) {
            $ValueName = $Matches[1]
        } else {
            $ValueName = '@'
        }

        $RegValue = $Line

        # Add any comments preceding the value
        if ($RegComment) {
            $RegValue = '{0}{1}{2}' -f $RegComment, [Environment]::NewLine, $RegValue
            $RegComment = [String]::Empty
        }

        # Is the value data multi-line?
        if ($Line -match $RegexMultilineValue) {
            do {
                $Idx++
                $ExtraLine = $FileContent[$Idx]
                $RegValue += '{0}{1}' -f [Environment]::NewLine, $ExtraLine
            } while ($ExtraLine -match '\\\s*$')
        }

        $SectionContent.Add($ValueName, $RegValue)
        continue
    }

    throw 'Unexpected content on line {0}: {1}' -f ($Idx + 1), $Line
}

# Add registry values for the final section
foreach ($ValueName in ($SectionContent.Keys | Sort-Object)) {
    $SortedContent.Add($SectionContent[$ValueName])
}

# Add any trailing comments
if ($RegComment) {
    $SortedContent.Add($RegComment)
    $RegComment = [String]::Empty
}

# Registry exports are UTF16-LE encoded
Set-Content -Path $RegFile -Value $SortedContent -Encoding 'unicode'
