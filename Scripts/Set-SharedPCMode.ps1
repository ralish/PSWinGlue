<#
    .SYNOPSIS
    Configures Shared PC Mode on Windows 10

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
    The MDM Bridge WMI Provider can only be interacted with from the NT AUTHORITY\SYSTEM account.

    Typically this script would be run non-interactively by a service running in the SYSTEM context (e.g. Group Policy Client).

    To run this script interactively you should use a tool like Sysinternals PsExec to run it under the SYSTEM account.

    For example, the following PsExec command will launch PowerShell under the SYSTEM account: psexec -s -i powershell

    Set up a shared or guest PC with Windows 10
    https://docs.microsoft.com/en-us/windows/configuration/set-up-shared-or-guest-pc

    SharedPC CSP
    https://docs.microsoft.com/en-us/windows/client-management/mdm/sharedpc-csp

    .LINK
    https://github.com/ralish/PSWinGlue
#>

#Requires -Version 3.0

[CmdletBinding()]
Param(
    [Bool]$EnableSharedPCMode,
    [Bool]$SetEduPolicies,
    [Bool]$SetPowerPolicies,
    [Int]$MaintenanceStartTime,
    [Bool]$SignInOnResume,
    [Int]$SleepTimeout,
    [Bool]$EnableAccountManager,
    [Int16]$AccountModel,
    [Int16]$DeletionPolicy,
    [Int16]$DiskLevelDeletion,
    [Int16]$DiskLevelCaching,
    [Bool]$RestrictLocalStorage,
    [String]$KioskModeAUMID,
    [String]$KioskModeUserTileDisplayText,
    [Int16]$InactiveThreshold,
    [Int16]$MaxPageFileSizeMB,

    [Switch]$PassThru
)

$OSBuild = [Environment]::OSVersion.Version.Build
$OSType = (Get-CimInstance -ClassName Win32_OperatingSystem).ProductType
if ($OSBuild -lt 14393 -or $OSType -ne 1) {
    throw 'Shared PC Mode is only available on Windows 10 1607 and newer.'
}

$SystemSid = 'S-1-5-18'
$CurrentSid = ([Security.Principal.WindowsIdentity]::GetCurrent()).User.Value
if ($CurrentSid -ne $SystemSid) {
    throw 'Must be running as SYSTEM to interact with MDM Bridge WMI Provider.'
}

$WmiNamespace = 'root\cimv2\mdm\dmmap'
$WmiClassName = 'MDM_SharedPC'
$MdmSharedPC = Get-CimInstance -Namespace $WmiNamespace -ClassName $WmiClassName -ErrorAction Stop

$MdmCspProperties = $MdmSharedPC.get_CimInstanceProperties().Name
$IgnoredParameters = [Management.Automation.Cmdlet]::CommonParameters + 'PassThru'

foreach ($Parameter in $PSBoundParameters.GetEnumerator()) {
    if ($Parameter.Key -notin $IgnoredParameters) {
        if ($Parameter.Key -notin $MdmCspProperties) {
            throw ('Parameter not supported on this Windows 10 version: {0}' -f $Parameter.Key)
        }

        $MdmSharedPC.($Parameter.Key) = $Parameter.Value
    }
}

Set-CimInstance -CimInstance $MdmSharedPC -ErrorAction Stop

if ($PassThru) {
    Get-CimInstance -Namespace $WmiNamespace -ClassName $WmiClassName
}
