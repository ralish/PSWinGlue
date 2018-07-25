[CmdletBinding()]
Param()

# Registry keys potentially containing Git installation details
$GitInstallRegKeys = @(
    # User: Native bitness
    'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Git_is1',
    # User: x86 on x64
    'HKCU:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Git_is1',
    # Machine: Native bitness
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Git_is1',
    # Machine: x86 on x64
    'HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Git_is1'
)

# Registry property which provides the installation directory
$GitInstallRegProperty = 'InstallLocation'

Function Initialize-GitEnvironment {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '')]
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)]
        [String]$Path
    )

    Write-Verbose -Message 'Initializing Git environment ...'

    # Amend the PATH variable to include the full set of Git utilities
    $Env:Path = '{0}\bin;{1}' -f $Path, $Env:Path

    # Create the HOME environment variable needed by SSH (via Git)
    #
    # Note that if running as a Scheduled Task the user profile of the
    # account we're running under may not yet be loaded on Windows 8 or
    # Server 2012 and newer. As a result, the USERPROFILE environment
    # variable will point to the Default user profile. We can work around
    # this by using GetFolderPath() with the UserProfile enumeration.
    #
    # See: https://support.microsoft.com/en-us/kb/2968540
    $Env:HOME = [Environment]::GetFolderPath([Environment+SpecialFolder]::UserProfile)
}

Function Test-GitInstalled {
    [CmdletBinding()]
    Param()

    Write-Verbose -Message 'Testing Git is installed ...'

    foreach ($RegKey in $GitInstallRegKeys) {
        $GitInstallPath = (Get-ItemProperty -Path $RegKey -ErrorAction Ignore).$GitInstallRegProperty

        if ($GitInstallPath) {
            break
        }
    }

    if (!$GitInstallPath) {
        throw 'Unable to locate a Git installation.'
    }

    if (!(Test-Path -Path $GitInstallPath -PathType Container)) {
        throw 'The Git installation directory does not exist.'
    }

    return $GitInstallPath
}

Function Test-GitRepository {
    [CmdletBinding()]
    Param()

    Write-Verbose -Message 'Testing current directory is part of a Git repository ...'

    $null = & git rev-parse --git-dir 2>&1
    if ($LASTEXITCODE -ne 0) {
        throw 'The current directory is not part of a Git repository.'
    }
}

Function Update-GitRepository {
    [CmdletBinding()]
    Param()

    Write-Verbose -Message 'Updating Git repository ...'

    & git add --all
    if ($LASTEXITCODE -ne 0) {
        throw 'Something went wrong adding all changes to the Git index.'
    }

    # Check if the index is dirty indicating we have something to commit
    & git diff-index --quiet --cached HEAD
    if ($LASTEXITCODE -ne 0) {
        $GitCommitDate = Get-Date -UFormat '%d/%m/%Y'
        git commit -m ('Changes for {0}' -f $GitCommitDate)
        if ($LASTEXITCODE -ne 0) {
            throw 'Something went wrong committing all changes in the Git index.'
        }
    }

    & git pull
    if ($LASTEXITCODE -ne 0) {
        throw 'Something went wrong pulling from the Git repository.'
    }

    & git push
    if ($LASTEXITCODE -ne 0) {
        throw 'Something went wrong pushing to the Git repository.'
    }
}

# Check Git is installed
$GitInstallPath = Test-GitInstalled

# Check we're operating within a Git repository
Test-GitRepository

# Initialize the Git environment
Initialize-GitEnvironment -Path $GitInstallPath

# Commit all changes and update the Git repository
Update-GitRepository
