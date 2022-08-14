<#
    .SYNOPSIS
    Uninstalls a specific font by name

    .DESCRIPTION
    Provides suport for uninstalling fonts by name.

    If the font is in-use and Administrator privileges are present, Windows will be configured to remove it next reboot.

    .PARAMETER Name
    The name of the font to be uninstalled, as per the font metadata.

    .PARAMETER Scope
    Specifies whether to uninstall the font from the system-wide or per-user fonts directory.

    Support for per-user fonts is only available from Windows 10 1809 and Windows Server 2019.

    Uninstalling system-wide fonts requires Administrator privileges.

    The default is system-wide.

    .EXAMPLE
    Uninstall-Font -Name 'Georgia (TrueType)'

    Uninstalls the "Georgia (TrueType)" font from the system-wide fonts directory.

    .NOTES
    Per-user fonts are only uninstalled in the context of the user executing the function.

    .LINK
    https://github.com/ralish/PSWinGlue
#>

#Requires -Version 3.0

[CmdletBinding(SupportsShouldProcess)]
[OutputType([Void])]
Param(
    [Parameter(Mandatory)]
    [String]$Name,

    [ValidateSet('System', 'User')]
    [String]$Scope = 'System'
)

$PowerShellCore = New-Object -TypeName Version -ArgumentList 6, 0
if ($PSVersionTable.PSVersion -ge $PowerShellCore -and $PSVersionTable.Platform -ne 'Win32NT') {
    throw '{0} is only compatible with Windows.' -f $MyInvocation.MyCommand.Name
}

Function Uninstall-Font {
    [CmdletBinding(SupportsShouldProcess)]
    [OutputType([Void])]
    Param(
        [Parameter(Mandatory)]
        [String]$Name,

        [ValidateSet('System', 'User')]
        [String]$Scope = 'System'
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
        $FontsReg = Get-Item -Path $FontsRegKey -ErrorAction Stop
    } catch {
        throw 'Unable to open {0} fonts registry key: {1}' -f $Scope.ToLower(), $FontsRegKey
    }

    if ($FontsReg.Property -notcontains $Name) {
        throw 'Font not registered for {0}: {1}' -f $Scope.ToLower(), $Name
    }

    $FontRegValue = $FontsReg.GetValue($Name)
    if ($Scope -eq 'User') {
        $FontFilePath = $FontRegValue
    } else {
        $FontFilePath = Join-Path -Path $FontsFolder -ChildPath $FontRegValue
    }

    if ($PSCmdlet.ShouldProcess($Name, 'Uninstall font')) {
        $RemoveOnReboot = $false

        try {
            Write-Debug -Message ('Removing font file: {0}' -f $FontFilePath)
            Remove-Item -Path $FontFilePath -ErrorAction Stop
        } catch [Management.Automation.ItemNotFoundException] {
            Write-Warning -Message ('Font file not found: {0}' -f $FontFilePath)
        } catch [UnauthorizedAccessException] {
            $RemoveOnReboot = $true
        }

        if ($RemoveOnReboot) {
            if (!(Test-IsAdministrator)) {
                throw 'Unable to uninstall in-use font. Retry as Administrator to remove on next reboot.'
            }

            if (!('PSWinGlue.UninstallFont' -as [Type])) {
                $MoveFileEx = @'
[DllImport("kernel32.dll", CharSet = CharSet.Unicode, EntryPoint = "MoveFileExW", ExactSpelling = true, SetLastError = true)]
public static extern bool MoveFileEx([MarshalAs(UnmanagedType.LPWStr)] string lpExistingFileName, IntPtr lpNewFileName, uint dwFlags);
'@

                Add-Type -Namespace 'PSWinGlue' -Name 'UninstallFont' -MemberDefinition $MoveFileEx
            }

            $MOVEFILE_DELAY_UNTIL_REBOOT = 4
            if ([PSWinGlue.UninstallFont]::MoveFileEx($FontFilePath, [IntPtr]::Zero, $MOVEFILE_DELAY_UNTIL_REBOOT) -eq $false) {
                throw (New-Object -TypeName 'ComponentModel.Win32Exception')
            }
        }

        Write-Debug -Message ('Removing font from {0} registry: {1}' -f $Scope.ToLower(), $Name)
        Remove-ItemProperty -Path $FontsRegKey -Name $Name
    }
}

Function Test-IsAdministrator {
    [CmdletBinding()]
    [OutputType([Boolean])]
    Param()

    $User = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
    if ($User.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        return $true
    }

    return $false
}

# Windows 10 1809 and Windows Server 2019 introduced support for installing
# fonts per-user. The corresponding Windows release build number is 17763.
Function Test-PerUserFontsSupported {
    [CmdletBinding()]
    [OutputType([Boolean])]
    Param()

    $BuildNumber = [Int](Get-CimInstance -ClassName 'Win32_OperatingSystem' -Verbose:$false).BuildNumber
    if ($BuildNumber -ge 17763) {
        return $true
    }

    return $false
}

if ($Scope -eq 'System' -and !(Test-IsAdministrator)) {
    throw 'Administrator privileges are required to uninstall system-wide fonts.'
} elseif ($Scope -eq 'User' -and !(Test-PerUserFontsSupported)) {
    throw 'Per-user fonts are only supported from Windows 10 1809 and Windows Server 2019.'
}

Uninstall-Font @PSBoundParameters
