#Requires -Version 3.0

[CmdletBinding()]
Param()

Function Initialize-GitEnvironment {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)]
        [String]$Path
    )

    # Amend the PATH variable to include the full set of Git utilities
    $Env:Path = '{0}\bin;{1}' -f $Path, $Env:Path

    # Create the HOME environment variable needed by SSH (via Git)
    #
    # If running as a Scheduled Task the user profile of the account may not yet be loaded on
    # Windows 8 or Server 2012 and newer. This results in the USERPROFILE environment variable
    # pointing to the Default user profile. We can work around this by using GetFolderPath().
    #
    # See: https://support.microsoft.com/en-us/kb/2968540
    $Env:HOME = [Environment]::GetFolderPath([Environment+SpecialFolder]::UserProfile)
}

Function Test-GitInstalled {
    [CmdletBinding()]
    Param()

    # Registry keys potentially containing Git installation details
    $GitRegistryKeys = @(
        # User
        'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Git_is1',
        # Machine: Native bitness
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Git_is1',
        # Machine: x86 on x64
        'HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Git_is1'
    )

    # Registry property which contains the installation directory
    $GitInstallProperty = 'InstallLocation'

    foreach ($RegKey in $GitRegistryKeys) {
        $GitInstallPath = (Get-ItemProperty -Path $RegKey -ErrorAction Ignore).$GitInstallProperty

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

    $null = & git rev-parse --git-dir 2>&1
    if ($LASTEXITCODE -ne 0) {
        throw ('The current directory does not belong to a Git repository: {0}' -f $PWD.Path)
    }
}

Function Update-GitRepository {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '')]
    [CmdletBinding()]
    Param()

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
        throw 'Something went wrong pulling from the remote Git repository.'
    }

    & git push
    if ($LASTEXITCODE -ne 0) {
        throw 'Something went wrong pushing to the remote Git repository.'
    }
}

# Test Git software is installed
$GitInstallPath = Test-GitInstalled

# Check we're in a Git repository
Test-GitRepository

# Initialize the Git environment
Initialize-GitEnvironment -Path $GitInstallPath

# Commit all and push all changes
Update-GitRepository
