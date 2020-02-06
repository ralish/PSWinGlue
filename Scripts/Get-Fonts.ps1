#Requires -Version 3.0

[CmdletBinding()]
Param(
    [ValidateSet('System', 'User')]
    [String]$Scope='System'
)

# Supported font extensions
$script:ValidExts = @('.otf', '.ttf')
$script:ValidExtsRegex = '\.(otf|ttf)$'

Function Get-Fonts {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '')]
    [CmdletBinding()]
    [OutputType([Object[]])]
    Param(
        [ValidateSet('System', 'User')]
        [String]$Scope='System'
    )

    switch ($Scope) {
        'System' {
            $FontsFolder = [Environment]::GetFolderPath('Fonts')
            $FontsRegKey = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts'
        }
        'User' {
            $FontsFolder = Join-Path -Path ([Environment]::GetFolderPath('LocalApplicationData')) -ChildPath 'Microsoft\Windows\Fonts'
            $FontsRegKey = 'HKCU:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts'
        }
    }

    try {
        $FontFiles = @(Get-ChildItem -Path $FontsFolder -ErrorAction Stop | Where-Object Extension -in $script:ValidExts)
    } catch {
        throw ('Unable to enumerate {0} fonts folder: {1}' -f $Scope.ToLower(), $FontsFolder)
    }

    try {
        $FontsReg = Get-Item -Path $FontsRegKey -ErrorAction Stop
    } catch {
        throw ('Unable to open {0} fonts registry key: {1}' -f $Scope.ToLower(), $FontsRegKey)
    }

    [Collections.ArrayList]$Fonts = @()
    [Collections.ArrayList]$FontsRegFileNames = @()
    foreach ($FontRegName in ($FontsReg.Property | Sort-Object)) {
        $FontRegValue = $FontsReg.GetValue($FontRegName)

        if ($Scope -eq 'User') {
            $FontRegFileName = [IO.Path]::GetFileName($FontRegValue)
        } else {
            $FontRegFileName = $FontRegValue
        }

        if ($FontRegFileName -notmatch $script:ValidExtsRegex) {
            Write-Debug -Message ('Ignoring font with unsupported extension: {0} -> {1}' -f $FontRegName, $FontRegFileName)
            continue
        } elseif ($FontFiles.Name -notcontains $FontRegFileName) {
            Write-Warning -Message ('Font file for registered font does not exist: {0} -> {1}' -f $FontRegName, $FontRegFileName)
            continue
        }

        $Font = [PSCustomObject]@{
            Name = $FontRegName
            File = $FontFiles | Where-Object Name -eq $FontRegFileName
        }

        $null = $Fonts.Add($Font)
        $null = $FontsRegFileNames.Add($FontRegFileName)
    }

    foreach ($FontFileName in $FontFiles.Name) {
        if ($FontFileName -notin $FontsRegFileNames) {
            Write-Warning -Message ('Font file not registered for {0}: {1}' -f $Scope.ToLower(), $FontFileName)
        }
    }

    return $Fonts
}

Function Test-PerUserFontsSupported {
    [CmdletBinding()]
    [OutputType([bool])]
    Param()

    # Windows 10 1809 introduced support for installing fonts per-user. The
    # corresponding release build number is 17763 (ignoring Insider builds).
    $BuildNumber = [Int](Get-CimInstance -ClassName 'Win32_OperatingSystem' -Verbose:$false).BuildNumber
    if ($BuildNumber -ge 17763) {
        return $true
    }

    return $false
}

if ($Scope -eq 'User' -and !(Test-PerUserFontsSupported)) {
    throw 'Per-user fonts are only supported from Windows 10 1809.'
}

Get-Fonts @PSBoundParameters
