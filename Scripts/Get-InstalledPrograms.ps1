<#
    .SYNOPSIS
    Retrieves installed programs

    .DESCRIPTION
    Enumerates all installed programs for the system and current user.

    The results should be nearly identical to those displayed via the "Programs and Features" view of the Windows Control Panel.

    For the "Apps & features" pane of the Settings app the results will be a subset of those displayed (see the Notes section).

    .EXAMPLE
    Get-InstalledPrograms

    Retrieves all programs installed system-wide or for the current user.

    .NOTES
    Only native Windows applications which register an uninstaller are displayed.

    Microsoft Store apps are not currently enumerated, which the "Apps & features" pane of the Settings app does display.

    The available information displayed for each program is expected to vary, as each program is itself responsible for recording it.

    If the installation date is not explicitly recorded by an installed program, we attempt to derive it based on the last write time of the registry key.

    There is no documented API for enumerating installed native Windows applications, so an approach based off reverse engineering Microsoft's implementation is used.

    There are three registry keys which are inspected to populate the list of installed programs:
    - System-wide in native bitness
      HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall
    - System-wide under the 32-bit emulation layer (64-bit Windows only)
      HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall
    - Current-user (any bitness)
      HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall

    .LINK
    https://github.com/ralish/PSWinGlue
#>

#Requires -Version 3.0

[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingEmptyCatchBlock', '')]
[CmdletBinding()]
[OutputType([PSCustomObject[]])]
Param()

$PowerShellCore = New-Object -TypeName Version -ArgumentList 6, 0
if ($PSVersionTable.PSVersion -ge $PowerShellCore -and $PSVersionTable.Platform -ne 'Win32NT') {
    throw '{0} is only compatible with Windows.' -f $MyInvocation.MyCommand.Name
}

if (!('PSWinGlue.GetInstalledPrograms' -as [Type])) {
    $RegQueryInfoKey = @'
[DllImport("advapi32.dll", EntryPoint = "RegQueryInfoKeyW")]
public static extern int RegQueryInfoKey(Microsoft.Win32.SafeHandles.SafeRegistryHandle hKey,
                                            IntPtr lpClass,
                                            IntPtr lpcchClass,
                                            IntPtr lpReserved,
                                            IntPtr lpcSubKeys,
                                            IntPtr lpcbMaxSubKeyLen,
                                            IntPtr lpcbMaxClassLen,
                                            IntPtr lpcValues,
                                            IntPtr lpcbMaxValueNameLen,
                                            IntPtr lpcbMaxValueLen,
                                            IntPtr lpcbSecurityDescriptor,
                                            out UInt64 lpftLastWriteTime);
'@

    $AddTypeParams = @{}

    if ($PSVersionTable['PSEdition'] -eq 'Core') {
        $AddTypeParams['ReferencedAssemblies'] = 'Microsoft.Win32.Registry'
    }

    Add-Type -Namespace 'PSWinGlue' -Name 'GetInstalledPrograms' -MemberDefinition $RegQueryInfoKey @AddTypeParams
}

$TypeName = 'PSWinGlue.InstalledProgram'
Update-TypeData -TypeName $TypeName -DefaultDisplayPropertySet @('Name', 'Publisher', 'Version', 'Scope') -Force

# System-wide in native bitness
$ComputerNativeRegPath = 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall'
# System-wide under the 32-bit emulation layer (64-bit Windows only)
$ComputerWow64RegPath = 'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
# Current-user (any bitness)
$UserRegPath = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall'

# Retrieve all installed programs from available keys
$UninstallKeys = Get-ChildItem -Path $ComputerNativeRegPath
if (Test-Path -Path $ComputerWow64RegPath -PathType Container) {
    $UninstallKeys += Get-ChildItem -Path $ComputerWow64RegPath
}
if (Test-Path -Path $UserRegPath -PathType Container) {
    $UninstallKeys += Get-ChildItem -Path $UserRegPath
}

# Filter out all the uninteresting installation results
$Results = New-Object -TypeName 'Collections.Generic.List[PSCustomObject]'
foreach ($UninstallKey in $UninstallKeys) {
    $Program = Get-ItemProperty -Path $UninstallKey.PSPath

    # Skip any program which doesn't define a display name
    if (!$Program.PSObject.Properties['DisplayName']) {
        continue
    }

    # Skip any program without an uninstall command which is not marked non-removable
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
        PSTypeName    = $TypeName
        PSPath        = $Program.PSPath
        Name          = $Program.DisplayName
        Publisher     = $null
        InstallDate   = $null
        EstimatedSize = $null
        Version       = $null
        Location      = $null
        Uninstall     = $null
        Scope         = $null
    }

    if ($Program.PSObject.Properties['Publisher']) {
        $Result.Publisher = $Program.Publisher
    }

    # Try and convert any InstallDate value to a DateTime
    if ($Program.PSObject.Properties['InstallDate']) {
        $RegInstallDate = $Program.InstallDate
        if ($RegInstallDate -match '^[0-9]{8}') {
            try {
                $Result.InstallDate = New-Object -TypeName 'DateTime' -ArgumentList $RegInstallDate.Substring(0, 4), $RegInstallDate.Substring(4, 2), $RegInstallDate.Substring(6, 2)
            } catch { }
        }

        if (!$Result.InstallDate) {
            Write-Warning -Message ('[{0}] Registry key has invalid value for InstallDate: {1}' -f $Program.DisplayName, $RegInstallDate)
        }
    }

    # Fall back to the last write time of the registry key
    if (!$Result.InstallDate) {
        [UInt64]$RegLastWriteTime = 0
        $Status = [PSWinGlue.GetInstalledPrograms]::RegQueryInfoKey($UninstallKey.Handle, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, [ref]$RegLastWriteTime)

        if ($Status -eq 0) {
            $Result.InstallDate = [DateTime]::FromFileTime($RegLastWriteTime)
        } else {
            Write-Warning -Message ('[{0}] Retrieving registry key last write time failed with status: {1}' -f $Program.DisplayName, $Status)
        }
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

    $Results.Add($Result)
}

return ($Results.ToArray() | Sort-Object -Property Name)
