#Requires -Version 3.0

[CmdletBinding()]
Param()

$Results = @()
$TypeName = 'PSWinGlue.InstalledProgram'

Update-TypeData -TypeName $TypeName -DefaultDisplayPropertySet @('Name', 'Publisher', 'Version', 'Scope') -Force

# Programs installed system-wide in native bitness
$ComputerNativeRegPath = 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall'
# Programs installed only under the current-user (any bitness)
$UserRegPath = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall'
# Programs installed system-wide under the 32-bit emulation layer (64-bit Windows only)
$ComputerWow64RegPath = 'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall'

# Retrieve all installed programs from available keys
$UninstallKeys = Get-ChildItem -Path $ComputerNativeRegPath
if (Test-Path -Path $ComputerWow64RegPath -PathType Container) {
    $UninstallKeys += Get-ChildItem -Path $ComputerWow64RegPath
}
$UninstallKeys += Get-ChildItem -Path $UserRegPath

# Filter out all the uninteresting installation results
foreach ($UninstallKey in $UninstallKeys) {
    $Program = Get-ItemProperty -Path $UninstallKey.PSPath

    # Skip any program which doesn't define a display name
    if (!$Program.PSObject.Properties['DisplayName']) {
        continue
    }

    # Ensure the program either:
    # - Has an uninstall command
    # - Is marked as non-removable
    if (!($Program.PSObject.Properties['UninstallString'] -or ($Program.PSObject.Properties['NoRemove'] -and $Program.NoRemove -eq 1))) {
        continue
    }

    # Skip any program which defines a parent program
    if ($Program.PSObject.Properties['ParentKeyName'] -or $Program.PSObject.Properties['ParentDisplayName']) {
        continue
    }

    # Skip any program marked as a system component
    if ($Program.PSObject.Properties['SystemComponent'] -and $Program.SystemComponent -eq 1) {
        continue
    }

    # Skip any program which defines a release type
    if ($Program.PSObject.Properties['ReleaseType']) {
        continue
    }

    $Result = [PSCustomObject]@{
        PSTypeName      = $TypeName
        PSPath          = $Program.PSPath
        Name            = $Program.DisplayName
        Publisher       = $null
        InstallDate     = $null
        EstimatedSize   = $null
        Version         = $null
        Location        = $null
        Uninstall       = $null
        Scope           = $null
    }

    if ($Program.PSObject.Properties['Publisher']) {
        $Result.Publisher = $Program.Publisher
    }

    if ($Program.PSObject.Properties['InstallDate']) {
        $Result.InstallDate = $Program.InstallDate
    }

    if ($Program.PSObject.Properties['EstimatedSize']) {
        $Result.EstimatedSize = $Program.EstimatedSize
    }

    if ($Program.PSObject.Properties['DisplayVersion']) {
        $Result.Version = $Program.DisplayVersion
    }

    if ($Program.PSObject.Properties['InstallLocation']) {
        $Result.Location = $Program.InstallLocation
    }

    if ($Program.PSObject.Properties['UninstallString']) {
        $Result.Uninstall = $Program.UninstallString
    }

    if ($Program.PSPath.startswith('Microsoft.PowerShell.Core\Registry::HKEY_LOCAL_MACHINE')) {
        $Result.Scope = 'System'
    } else {
        $Result.Scope = 'User'
    }

    $Results += $Result
}

return ($Results | Sort-Object -Property Name)
