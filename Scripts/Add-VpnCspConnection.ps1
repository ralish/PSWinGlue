#Requires -Version 3.0

[CmdletBinding()]
Param(
    [Parameter(Mandatory)]
    [String]$ProfileName,

    [Parameter(ParameterSetName = 'Content', Mandatory)]
    [String]$ProfileXmlContent,

    [Parameter(ParameterSetName = 'File', Mandatory)]
    [String]$ProfileXmlFile,

    [Switch]$PassThru
)

$OSBuild = [Environment]::OSVersion.Version.Build
$OSType = (Get-CimInstance -ClassName Win32_OperatingSystem).ProductType
if ($OSBuild -lt 14393 -or $OSType -ne 1) {
    throw 'VPN configuration with ProfileXML is only available on Windows 10 1607 and newer.'
}

$SystemSid = 'S-1-5-18'
$CurrentSid = ([Security.Principal.WindowsIdentity]::GetCurrent()).User.Value
if ($CurrentSid -ne $SystemSid) {
    throw 'Must be running as SYSTEM to interact with MDM Bridge WMI Provider.'
}

if ($PSCmdlet.ParameterSetName -eq 'File') {
    $ProfileXmlContent = Get-Content -Path $ProfileXmlFile -Raw -ErrorAction Stop
}

try {
    $ProfileXml = [Xml]$ProfileXmlContent
} catch {
    throw $_
}

$WmiNamespace = 'root\cimv2\mdm\dmmap'
$WmiClassName = 'MDM_VPNv2_01'
$MdmCspPath = './Vendor/MSFT/VPNv2'
$MdmCspProperties = @{
    ParentID   = $MdmCspPath
    InstanceID = [Uri]::EscapeDataString($ProfileName)
    ProfileXML = [Security.SecurityElement]::Escape($ProfileXml.InnerXml)
}

$VpnProfile = New-CimInstance -Namespace $WmiNamespace -ClassName $WmiClassName -Property $MdmCspProperties -ErrorAction Stop

if ($PassThru) {
    return $VpnProfile
}
