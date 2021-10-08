#Requires -Version 3.0

[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingEmptyCatchBlock', '')]
[CmdletBinding()]
Param()

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

if (!('PSWinGlue.GetInstalledPrograms' -as [Type])) {
    $AddTypeParams = @{}

    if ($PSVersionTable['PSEdition'] -eq 'Core') {
        $AddTypeParams['ReferencedAssemblies'] = 'Microsoft.Win32.Registry'
    }

    Add-Type -Namespace PSWinGlue -Name GetInstalledPrograms -MemberDefinition $RegQueryInfoKey @AddTypeParams
}

$Results = New-Object -TypeName Collections.ArrayList
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
if (Test-Path -Path $UserRegPath -PathType Container) {
    $UninstallKeys += Get-ChildItem -Path $UserRegPath
}

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

    # Try and convert the InstallDate value to a DateTime
    if ($Program.PSObject.Properties['InstallDate']) {
        $RegInstallDate = $Program.InstallDate
        if ($RegInstallDate -match '^[0-9]{8}') {
            try {
                $Result.InstallDate = New-Object -TypeName DateTime -ArgumentList $RegInstallDate.Substring(0, 4), $RegInstallDate.Substring(4, 2), $RegInstallDate.Substring(6, 2)
            } catch { }
        }

        if (!$Result.InstallDate) {
            Write-Warning -Message ('[{0}] Registry key has invalid value for InstallDate: {1}' -f $Program.DisplayName, $RegInstallDate)
        }
    }

    # Fall back to the last write time of the registry key
    if (!$Result.InstallDate) {
        [UInt64]$RegLastWriteTime = 0
        $Status = [PSWinGlue.GetInstalledPrograms]::RegQueryInfoKey($UninstallKey.Handle, [IntPtr]::Zero, [IntPtr]::Zero, [IntPtr]::Zero, [IntPtr]::Zero, [IntPtr]::Zero, [IntPtr]::Zero, [IntPtr]::Zero, [IntPtr]::Zero, [IntPtr]::Zero, [IntPtr]::Zero, [ref]$RegLastWriteTime)

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

    $null = $Results.Add($Result)
}

return ($Results | Sort-Object -Property Name)
