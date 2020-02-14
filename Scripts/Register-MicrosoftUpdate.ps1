#Requires -Version 3.0
#Requires -RunAsAdministrator

[CmdletBinding()]
Param()

# https://docs.microsoft.com/en-us/windows/desktop/wua_sdk/opt-in-to-microsoft-update

# Microsoft Update identifier
$ServiceID = '7971f918-a847-4430-9279-4a52d1efe18d'
# AddServiceFlag enumeration
# https://docs.microsoft.com/en-au/windows/desktop/api/wuapi/ne-wuapi-tagaddserviceflag
$ServiceFlags = 7

try {
    $ServiceManager = New-Object -ComObject Microsoft.Update.ServiceManager
} catch {
    throw $_
}

try {
    $ServiceRegistration = $ServiceManager.AddService2($ServiceID, $ServiceFlags, '')
    $null = [Runtime.InteropServices.Marshal]::FinalReleaseComObject($ServiceRegistration)
} catch {
    throw $_
} finally {
    $null = [Runtime.InteropServices.Marshal]::FinalReleaseComObject($ServiceManager)
}
