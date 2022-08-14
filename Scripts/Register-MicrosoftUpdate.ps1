<#
    .SYNOPSIS
    Register the Microsoft Update service with the Windows Update Agent

    .DESCRIPTION
    The Windows Update Agent (WUA) has a concept of "services" which, once registered with WUA, add additional catalogues of updates for delivery to the system.

    By default, the Windows Update service is registered to provide updates to the operating system. To receive updates for other installed Microsoft products the Microsoft Update service must be registered.

    This script automates the registration of the Microsoft Update service with WUA to ensure the operating system and all installed Microsoft products receive available updates.

    .EXAMPLE
    Register-MicrosoftUpdate

    Ensures the Microsoft Update service is registered with the Windows Update Agent.

    .NOTES
    Administrator privileges are required to modify the Windows Update Agent configuration.

    .LINK
    https://github.com/ralish/PSWinGlue
#>

#Requires -Version 3.0

[CmdletBinding()]
[OutputType()]
Param()

$PowerShellCore = New-Object -TypeName Version -ArgumentList 6, 0
if ($PSVersionTable.PSVersion -ge $PowerShellCore -and $PSVersionTable.Platform -ne 'Win32NT') {
    throw '{0} is only compatible with Windows.' -f $MyInvocation.MyCommand.Name
}

$User = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
if (!$User.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    throw '{0} requires Administrator privileges.' -f $MyInvocation.MyCommand.Name
}

# Opt-In to Microsoft Update
# https://docs.microsoft.com/en-us/windows/win32/wua_sdk/opt-in-to-microsoft-update
$MuServiceId = '7971f918-a847-4430-9279-4a52d1efe18d'

# AddServiceFlag enumeration
# https://docs.microsoft.com/en-us/windows/win32/api/wuapi/ne-wuapi-addserviceflag
$asfAllowPendingRegistration = 1
$asfAllowOnlineRegistration = 2
$asfRegisterServiceWithAU = 4
$MuServiceFlags = $asfAllowPendingRegistration + $asfAllowOnlineRegistration + $asfRegisterServiceWithAU

$WuaServiceManager = $null
$WuaServices = $null
$MuService = $null

try {
    $WuaServiceManager = New-Object -ComObject 'Microsoft.Update.ServiceManager'

    $MuServiceRegistered = $false
    $WuaServices = $WuaServiceManager.Services
    for ($i = 0; $i -lt $WuaServices.Count; $i++) {
        $WuaService = $WuaServices[$i]

        if ($WuaService.ServiceID -contains $MuServiceId) {
            $MuServiceRegistered = $true
        }

        $null = [Runtime.InteropServices.Marshal]::ReleaseComObject($WuaService)

        if ($MuServiceRegistered) {
            Write-Verbose -Message 'Microsoft Update service already registered with WUA.'
            return
        }
    }

    Write-Verbose -Message 'Registering Microsoft Update service with WUA ...'
    $MuService = $WuaServiceManager.AddService2($MuServiceId, $MuServiceFlags, '')
} catch {
    throw $_
} finally {
    if ($MuService) { $null = [Runtime.InteropServices.Marshal]::ReleaseComObject($MuService) }
    if ($WuaServices) { $null = [Runtime.InteropServices.Marshal]::ReleaseComObject($WuaServices) }
    if ($WuaServiceManager) { $null = [Runtime.InteropServices.Marshal]::ReleaseComObject($WuaServiceManager) }
}
