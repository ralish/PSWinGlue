<#
    .SYNOPSIS
    Adds a VPN connection using the VPNv2 CSP via the MDM Bridge WMI Provider

    .DESCRIPTION
    Uses the MDM Bridge WMI Provider to interact with the MDM_VPNv2_01 class for adding VPNv2 connections.

    .PARAMETER ProfileName
    The profile name for the VPN connection.

    .PARAMETER ProfileXml
    The ProfileXML specifying the settings for the VPN connection.

    .PARAMETER ProfilePath
    The path to the file containing the ProfileXML specifying the settings for the VPN connection.

    .PARAMETER PassThru
    Return the created WMI instance corresponding to the VPN profile.

    .EXAMPLE
    Add-VpnCspConnection -ProfileName 'My VPN' -ProfilePath 'D:\My VPN.xml'

    Creates a new VPN profile named "My VPN" using the ProfileXML in the "D:\My VPN.xml" file.

    .NOTES
    VPN configuration using the VPNv2 CSP is only available on Windows 10 1607 or later.

    To interact with the MDM Bridge WMI Provider the function must be running as SYSTEM.

    Typically this function would be run non-interactively by a service running in the SYSTEM context (e.g. Group Policy Client).

    To run this function interactively you should use a tool like Sysinternals PsExec to run it under the SYSTEM account.

    For example, the following PsExec command will launch PowerShell under the SYSTEM account: psexec -s -i powershell

    VPNv2 CSP
    https://docs.microsoft.com/en-us/windows/client-management/mdm/vpnv2-csp

    MDM_VPNv2_01 class
    https://docs.microsoft.com/en-us/windows/win32/dmwmibridgeprov/mdm-vpnv2-01

    .LINK
    https://github.com/ralish/PSWinGlue
#>

# Minimum supported Windows release ships with PowerShell 5.1
#Requires -Version 5.1

[CmdletBinding()]
[OutputType([Void], [Microsoft.Management.Infrastructure.CimInstance])]
Param(
    [Parameter(Mandatory)]
    [String]$ProfileName,

    [Parameter(ParameterSetName = 'Path', Mandatory)]
    [String]$ProfilePath,

    [Parameter(ParameterSetName = 'Xml', Mandatory)]
    [Xml]$ProfileXml,

    [Switch]$PassThru
)

$WmiNamespace = 'root\cimv2\mdm\dmmap'
$WmiClassName = 'MDM_VPNv2_01'
$MdmCspPath = './Vendor/MSFT/VPNv2'

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
    throw 'VPN configuration with ProfileXML is only available on Windows 10 1607 or later.'
}

$PowerShellMin = New-Object -TypeName Version -ArgumentList 5, 1
if ($PSVersionTable.PSVersion -lt $PowerShellMin) {
    throw '{0} requires at least PowerShell {1}.' -f $MyInvocation.MyCommand.Name, $PowerShellMin
}

$SidCurrent = ([Security.Principal.WindowsIdentity]::GetCurrent()).User.Value
if ($SidCurrent -ne $SidSystem) {
    throw 'Must be running as SYSTEM to interact with MDM Bridge WMI Provider.'
}

if ($PSCmdlet.ParameterSetName -eq 'Path') {
    try {
        $ProfileXml = [Xml](Get-Content -Path $ProfilePath -Raw -ErrorAction Stop)
    } catch {
        throw $_
    }
}

$MdmCspProperties = @{
    ParentID   = $MdmCspPath
    InstanceID = [Uri]::EscapeDataString($ProfileName)
    ProfileXML = [Security.SecurityElement]::Escape($ProfileXml.InnerXml)
}

$VpnProfile = New-CimInstance -Namespace $WmiNamespace -ClassName $WmiClassName -Property $MdmCspProperties -ErrorAction Stop

if ($PassThru) {
    return $VpnProfile
}
