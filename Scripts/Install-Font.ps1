[CmdletBinding(SupportsShouldProcess)]
Param(
    [Parameter(Mandatory)]
    [String]$Path,

    [Switch]$Recurse
)

# Valid font extensions
$ValidExts = @('.fon', '.otf', '.ttc', '.ttf')

try {
    $FontPath = Get-Item -Path $Path
} catch {
    throw ('Provided path is invalid: {0}' -f $Path)
}

$Fonts = @()
if ($FontPath -is [IO.DirectoryInfo]) {
    $Fonts += Get-ChildItem -Path $FontPath -Recurse:$Recurse | Where-Object -Property Extension -in $ValidExts

    if (!$Fonts) {
        throw ('Unable to locate any fonts in provided directory: {0}' -f $FontPath.FullName)
    }
} elseif ($FontPath -is [IO.FileInfo]) {
    if ($FontPath.Extension -notin $ValidExts) {
        throw ('Provided file does not appear to be a valid font: {0}' -f $FontPath.FullName)
    }

    $Fonts += $FontPath
} else {
    throw ('Expected directory or file but received: {0}' -f $FontPath.GetType().Name)
}

# ShellSpecialFolderConstants enumeration
# https://docs.microsoft.com/en-us/windows/desktop/api/Shldisp/ne-shldisp-shellspecialfolderconstants
$ssfFONTS = 20

# _SHFILEOPSTRUCTA structure
# https://docs.microsoft.com/en-us/windows/desktop/api/shellapi/ns-shellapi-_shfileopstructa
$FOF_SILENT = 4
$FOF_NOCONFIRMATION = 16
$FOF_NOERRORUI = 1024
$FOF_NOCOPYSECURITYATTRIBS = 2048

$ShellApp = New-Object -ComObject Shell.Application
$FontsFolder = $ShellApp.NameSpace($ssfFONTS)
$CopyOptions = $FOF_SILENT + $FOF_NOCONFIRMATION + $FOF_NOERRORUI + $FOF_NOCOPYSECURITYATTRIBS

foreach ($Font in $Fonts) {
    if ($PSCmdlet.ShouldProcess($Font.BaseName, 'Install font')) {
        $FontsFolder.CopyHere($Font.FullName, $CopyOptions)
    }
}
