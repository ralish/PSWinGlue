<#
    .SYNOPSIS
    Updates a Git repository

    .DESCRIPTION
    A helper function to automate updating a Git repository.

    The following steps are performed with a failure aborting subsequent steps:
    - Check the current directory is a valid Git repository
    - Add all changes in the working tree to the staging area
    - If there are staged changes commit them to the repository
    - Pull changes from upstream and perform a fast-forward merge
    - Push the new commit(s) to the upstream repository

    .EXAMPLE
    Update-GitRepository

    Updates the Git repository in the current directory. See the description section for the steps performed.

    .NOTES
    This script is primarily intended for usage in unattended scenarios, such as invoking from a Scheduled Task.

    The following potential issues are handled, which are more common with or specific to Scheduled Task execution:
    - Checks Git is installed, and if so, retrieves its installation directory
      This enables the script to handle the case where the Git binaries aren't in the PATH of the executing user.
    - Ensures the environment is initialised for Git to function correctly
      In particular, the HOME environment variable may be missing when run as a Scheduled Task, causing SSH issues.

    .LINK
    https://github.com/ralish/PSWinGlue
#>

#Requires -Version 3.0

[CmdletBinding()]
[OutputType([String[]])]
Param()

$PowerShellCore = New-Object -TypeName Version -ArgumentList 6, 0
if ($PSVersionTable.PSVersion -ge $PowerShellCore -and $PSVersionTable.Platform -ne 'Win32NT') {
    throw '{0} is only compatible with Windows.' -f $MyInvocation.MyCommand.Name
}

Function Initialize-GitEnvironment {
    [CmdletBinding()]
    [OutputType([Void])]
    Param(
        [Parameter(Mandatory)]
        [String]$Path
    )

    # Amend the PATH variable to include the full set of Git utilities
    $Env:Path = '{0}\bin;{1}' -f $Path, $Env:Path

    # Create the HOME environment variable needed by SSH (via Git)
    #
    # When running as a Scheduled Task the user profile of the account may not
    # yet be loaded on Windows 8 or Server 2012 and newer. This can cause the
    # USERPROFILE environment variable to point to the Default user profile. A
    # workaround is to retrieve and set the correct path using GetFolderPath().
    #
    # Scheduled tasks reference incorrect user profile paths
    # https://support.microsoft.com/en-us/kb/2968540
    $Env:HOME = [Environment]::GetFolderPath([Environment+SpecialFolder]::UserProfile)
}

Function Test-GitInstalled {
    [CmdletBinding()]
    [OutputType([String])]
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
        try {
            $RegKeyProps = Get-ItemProperty -Path $RegKey -ErrorAction Stop
            $GitInstallPath = $RegKeyProps.$GitInstallProperty
        } catch {
            continue
        }

        if ([String]::IsNullOrWhiteSpace($GitInstallPath)) {
            continue
        }

        if (!(Test-Path -Path $GitInstallPath -PathType Container)) {
            continue
        }

        return $GitInstallPath
    }

    throw 'Unable to locate a Git installation.'
}

Function Test-GitRepository {
    [CmdletBinding()]
    [OutputType([Void])]
    Param()

    $null = & git rev-parse --git-dir 2>&1
    if ($LASTEXITCODE -ne 0) {
        throw 'The current directory does not belong to a Git repository: {0}' -f $PWD.Path
    }
}

Function Update-GitRepository {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '')]
    [CmdletBinding()]
    [OutputType([String[]])]
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
