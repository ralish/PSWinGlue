<#
    .SYNOPSIS
    Helper script to update OneDrive during Windows image creation

    .DESCRIPTION
    This script is designed to be used during creation of Windows images for deployment to client systems.

    The version of OneDrive bundled with a Windows release will often be out-of-date by the time of deployment to a client system. While OneDrive will update itself, some updates require Administrative privileges, which users may not have. By pre-installing the latest update we can hopefully avoid unwanted UAC prompts, while also streamlining the initial logon process.

    .PARAMETER SetupDestDir
    Path to the directory where we'll store a copy of the OneDrive setup file.

    The default is $env:ProgramData\Microsoft\OneDriveSetup.

    .PARAMETER SetupFilePath
    Path to the OneDrive setup file.

    The default is OneDriveSetup.exe in the current working directory.

    .PARAMETER SkipRunningSetup
    Skips executing OneDrive setup.

    .PARAMETER SkipUpdatingDefaultProfile
    Skips updating the default user profile to automatically run OneDrive setup on login.

    .EXAMPLE
    Update-OneDriveSetup

    Copies OneDrive setup to the default destination path, runs setup and updates the default user profile.

    .NOTES
    OneDrive release notes
    https://support.office.com/en-us/article/onedrive-release-notes-845dcf18-f921-435e-bf28-4e24b95e5fc0

    .LINK
    https://github.com/ralish/PSWinGlue
#>

#Requires -Version 3.0
#Requires -RunAsAdministrator

[CmdletBinding()]
Param(
    [ValidateNotNullOrEmpty()]
    [String]$SetupFilePath = 'OneDriveSetup.exe',

    [ValidateNotNullOrEmpty()]
    [String]$SetupDestDir = "$env:ProgramData\Microsoft\OneDriveSetup",

    [Switch]$SkipRunningSetup,
    [Switch]$SkipUpdatingDefaultProfile
)

$GetDefaultUserProfileDirectory = @'
[DllImport("userenv.dll", EntryPoint = "GetDefaultUserProfileDirectoryW", SetLastError = true)]
public static extern bool GetDefaultUserProfileDirectory(IntPtr lpProfileDir, out uint lpcchSize);

[DllImport("userenv.dll", CharSet = CharSet.Unicode, EntryPoint = "GetDefaultUserProfileDirectoryW", ExactSpelling = true, SetLastError = true)]
public static extern bool GetDefaultUserProfileDirectory(System.Text.StringBuilder lpProfileDir, out uint lpcchSize);
'@

if ($SkipRunningSetup -and $SkipUpdatingDefaultProfile) {
    throw 'Nothing to do as all operations skipped!'
}

$SetupFile = Get-Item -Path $SetupFilePath -ErrorAction Stop
if ($SetupFile -isnot [IO.FileInfo]) {
    throw 'OneDrive setup path is not a file: {0}' -f $SetupFilePath
} elseif ($SetupFile.Extension -ne '.exe') {
    throw 'OneDrive setup file must be an executable: {0}' -f $SetupFilePath
}

if (!(Test-Path -Path $SetupDestDir -PathType Container)) {
    Write-Verbose -Message 'Creating OneDrive setup directory ...'
    $null = New-Item -Path $SetupDestDir -ItemType Directory -Force -ErrorAction Ignore
}

$SetupDest = Get-Item -Path $SetupDestDir -ErrorAction Stop
if ($SetupDest -isnot [IO.DirectoryInfo]) {
    throw 'OneDrive setup destination is not a directory: {0}' -f $SetupDestDir
}

Write-Verbose -Message 'Copying OneDrive setup file ...'
$SetupDestFilePath = Join-Path -Path $SetupDestDir -ChildPath 'OneDriveSetup.exe'
Copy-Item -Path $SetupFilePath -Destination $SetupDestFilePath -Force
[IO.FileInfo]$SetupFilePath = Get-Item -Path $SetupDestFilePath -ErrorAction Stop

if (!$SkipRunningSetup) {
    Write-Verbose -Message 'Terminating any existing OneDrive setup ...'
    Stop-Process -Name 'OneDriveSetup' -ErrorAction Ignore

    Write-Verbose -Message 'Running OneDrive setup ...'
    $OneDriveSetup = Start-Process -FilePath $SetupFilePath.FullName -ArgumentList @('/silent') -Wait -PassThru
    if ($OneDriveSetup.ExitCode -notin ('0', '3010')) {
        Write-Warning -Message ('OneDrive setup returned exit code: {0}' -f $OneDriveSetup.ExitCode)
    }
}

if (!$SkipUpdatingDefaultProfile) {
    Write-Verbose -Message 'Updating OneDrive setup path for default user profile ...'

    if (!('PSWinGlue.UpdateOneDriveSetup' -as [Type])) {
        Add-Type -Namespace PSWinGlue -Name UpdateOneDriveSetup -MemberDefinition $GetDefaultUserProfileDirectory
    }

    $DefaultUserProfileBufSize = 0
    $null = [PSWinGlue.UpdateOneDriveSetup]::GetDefaultUserProfileDirectory([IntPtr]::Zero, [ref]$DefaultUserProfileBufSize)
    if (!$DefaultUserProfileBufSize) {
        throw 'Failed to determine buffer size for GetDefaultUserProfileDirectory().'
    }

    $DefaultUserProfilePath = New-Object -TypeName Text.StringBuilder -ArgumentList ($DefaultUserProfileBufSize - 1)
    if ([PSWinGlue.UpdateOneDriveSetup]::GetDefaultUserProfileDirectory($DefaultUserProfilePath, [ref]$DefaultUserProfileBufSize) -eq $false) {
        throw (New-Object -TypeName ComponentModel.Win32Exception)
    }

    $DefaultUserProfileHive = Join-Path -Path $DefaultUserProfilePath.ToString() -ChildPath 'NTUSER.DAT'
    Write-Debug -Message ('Default user profile registry hive: {0}' -f $DefaultUserProfileHive)

    $Registry = Start-Process -FilePath reg -ArgumentList @('LOAD', 'HKLM\OneDriveSetup', "`"$DefaultUserProfileHive`"") -Wait -PassThru
    if ($Registry.ExitCode -ne 0) {
        throw 'Failed to load the default user profile registry hive with error code: {0}' -f $Registry.ExitCode
    }

    Set-ItemProperty -Path 'HKLM:\OneDriveSetup\Software\Microsoft\Windows\CurrentVersion\Run' -Name 'OneDriveSetup' -Value $SetupFilePath.FullName

    $Registry = Start-Process -FilePath reg -ArgumentList @('UNLOAD', 'HKLM\OneDriveSetup') -Wait -PassThru
    if ($Registry.ExitCode -ne 0) {
        throw 'Failed to unload the default user profile registry hive with error code: {0}' -f $Registry.ExitCode
    }
}
