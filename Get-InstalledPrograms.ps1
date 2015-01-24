Function Get-InstalledPrograms {
    <#
        .Synopsis
        Fetches the list of installed software on a system via the Windows Registry.
        .Description
        Returns a list of software installed on a system determined from installations that have
        registered themselves in the Windows Registry. This cmdlet will parse both the native key
        and the WOW64 key if it exists to ensure a complete list of software installs is returned.
        .Notes
        This cmdlet is particularly useful on Server Core installations where the Programs and
        Features Control Panel applet isn't available and no equivalent cmdlet functionality exists.
    #>

    $ErrorActionPreference = 'Stop'

    # Define the key registry paths we'll retrieve installs from
    $NativeRegPath = 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*'
    $Wow6432RegPath = 'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'

    # Get the list of installed programs including WOW64 if present
    $UninstKeys = Get-ItemProperty $NativeRegPath
    if (Test-Path $Wow6432RegPath -PathType Container) {
        $UninstKeys += Get-ItemProperty $Wow6432RegPath
    }

    # Parse all returned installs and add them to an array
    $InstProgs = @()
    foreach  ($Prog in $UninstKeys) {
        # If the entry has no defined DisplayName ignore it as it's probably not useful
        if ($Prog.DisplayName -ne $null) {
            $ProgInfo = [PsCustomObject]@{
                Name = $Prog.DisplayName
                Publisher = $Prog.Publisher
                InstalledOn = $Prog.InstallDate
                Size = $Prog.EstimatedSize
                Version = $Prog.DisplayVersion
                Location = $Prog.InstallLocation
                Uninstall = $Prog.UninstallString
            }
            $ProgInfo.PSTypeNames.Add('PSWinGlue.Programs')
            $InstProgs += $ProgInfo
        }
    }
    return $InstProgs
}
