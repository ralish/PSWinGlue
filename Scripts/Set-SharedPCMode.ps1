<#
    .SYNOPSIS
    Configures Shared PC Mode using the SharedPC CSP via the MDM Bridge WMI Provider

    .DESCRIPTION
    The typical approach for configuring Shared PC Mode is to use an MDM solution which interacts with the SharedPC CSP.

    Alternatively, configuration can be performed by using the MDM Bridge WMI Provider to interact with the CSP via WMI.

    This script eases configuration when using the latter approach by providing a simple command to set available settings.

    For details on parameters consult the Shared PC Mode documentation (see the NOTES section of this command for links).

    .PARAMETER PassThru
    Return an instance of the MDM_SharedPC class after applying the requested configuration.

    .EXAMPLE
    Set-SharedPCMode -EnableSharedPCMode $true -EnableAccountManager $true -RestrictLocalStorage $true

    Enables Shared PC Mode with automatic account management and local storage restrictions.

    .NOTES
    VPN configuration using the VPNv2 CSP is only available on Windows 10 1607 or later.

    To interact with the MDM Bridge WMI Provider the script must be running as SYSTEM.

    Typically this script would be run non-interactively by a service running in the SYSTEM context (e.g. Group Policy Client).

    To run this script interactively you should use a tool like Sysinternals PsExec to run it under the SYSTEM account.

    For example, the following PsExec command will launch PowerShell under the SYSTEM account: psexec -s -i powershell

    Set up a shared or guest PC with Windows 10/11
    https://docs.microsoft.com/en-us/windows/configuration/set-up-shared-or-guest-pc

    SharedPC CSP
    https://docs.microsoft.com/en-us/windows/client-management/mdm/sharedpc-csp

    MDM_SharedPC class
    https://docs.microsoft.com/en-us/windows/win32/dmwmibridgeprov/mdm-sharedpc

    .LINK
    https://github.com/ralish/PSWinGlue
#>

# Minimum supported Windows release ships with PowerShell 5.1
#Requires -Version 5.1

[CmdletBinding()]
[OutputType([Void], [Microsoft.Management.Infrastructure.CimInstance])]
Param(
    [Bool]$EnableSharedPCMode,
    [Bool]$SetEduPolicies,
    [Bool]$SetPowerPolicies,
    [Int]$MaintenanceStartTime,
    [Bool]$SignInOnResume,
    [Int]$SleepTimeout,
    [Bool]$EnableAccountManager,
    [Int]$AccountModel,
    [Int]$DeletionPolicy,
    [Int]$DiskLevelDeletion,
    [Int]$DiskLevelCaching,
    [Bool]$RestrictLocalStorage,
    [String]$KioskModeAUMID,
    [String]$KioskModeUserTileDisplayText,
    [Int]$InactiveThreshold,
    [Int]$MaxPageFileSizeMB,

    [Switch]$PassThru
)

$WmiNamespace = 'root\cimv2\mdm\dmmap'
$WmiClassName = 'MDM_SharedPC'

$OSRequiredType = 1         # Workstation
$OSRequiredBuild = 14393    # Windows 10 1607
$SidSystem = 'S-1-5-18'     # NT AUTHORITY\SYSTEM

$PowerShellCore = New-Object -TypeName Version -ArgumentList 6, 0
if ($PSVersionTable.PSVersion -ge $PowerShellCore -and $PSVersionTable.Platform -ne 'Win32NT') {
    throw '{0} is only compatible with Windows.' -f $MyInvocation.MyCommand.Name
}

$OSCurrentType = (Get-CimInstance -ClassName 'Win32_OperatingSystem' -Verbose:$false).ProductType
$OSCurrentBuild = [Environment]::OSVersion.Version.Build
if ($OSCurrentBuild -lt $OSRequiredBuild -or $OSCurrentType -ne $OSRequiredType) {
    throw 'Shared PC Mode is only available on Windows 10 1607 or later.'
}

$SidCurrent = ([Security.Principal.WindowsIdentity]::GetCurrent()).User.Value
if ($SidCurrent -ne $SidSystem) {
    throw 'Must be running as SYSTEM to interact with MDM Bridge WMI Provider.'
}

$MdmSharedPC = Get-CimInstance -Namespace $WmiNamespace -ClassName $WmiClassName -ErrorAction Stop
$MdmCspProperties = $MdmSharedPC.get_CimInstanceProperties().Name
$IgnoredParameters = [Management.Automation.Cmdlet]::CommonParameters + 'PassThru'

foreach ($Parameter in $PSBoundParameters.GetEnumerator()) {
    if ($Parameter.Key -notin $IgnoredParameters) {
        if ($Parameter.Key -notin $MdmCspProperties) {
            throw 'Parameter not supported on this Windows 10 version: {0}' -f $Parameter.Key
        }

        $MdmSharedPC.($Parameter.Key) = $Parameter.Value
    }
}

Set-CimInstance -CimInstance $MdmSharedPC -ErrorAction Stop

if ($PassThru) {
    Get-CimInstance -Namespace $WmiNamespace -ClassName $WmiClassName
}
