Function Update-GitRepository {
    [CmdletBinding()]

    # The path to the Registry key containing the Git installation details
    $GitInstallRegPath = 'HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Git_is1'
    # The name of the Registry property that gives the installation location
    $GitInstallDirProp = 'InstallLocation'

    # Ensure that any errors we receive are considered fatal
    $ErrorActionPreference = 'Stop'

    Function Test-GitInstalled {
        Write-Verbose 'Testing Git is installed...'

        if (!(Test-Path $GitInstallRegPath -PathType Container)) {
            Write-Error 'Git does not appear to be installed on this system.'
        } else {
            $GitPath = (Get-ItemProperty -Path $GitInstallRegPath).$GitInstallDirProp
            if (!(Test-Path $GitPath -PathType Container)) {
                Write-Error 'The Git installation on this system appears to be damaged.'
            }
        }

        # Amend the PATH variable to include the full set of Git utilities
        $Env:Path="$Env:Path;$GitPath\bin"

        # Setup the HOME environment variable needed by SSH (this caused some serious pain)
        $Env:HOME = $Env:USERPROFILE
    }

    Function Test-GitRepository {
        Write-Verbose 'Testing current directory is a Git repository...'

        git rev-parse --git-dir 2>&1
        if ($LASTEXITCODE -ne 0) {
            Write-Error 'The current directory is not part of a Git repository.'
        }
    }

    Function Test-Windows64bit {
        Write-Verbose 'Testing if we are running on 64-bit Windows...'

        if ((Get-WmiObject 'Win32_OperatingSystem').OSArchitecture -ne '64-bit') {
            Write-Error 'We only support running on 64-bit systems. Seriously, it is time to upgrade already!'
        }
    }

    Function Update-GitRepository {
        Write-Verbose 'Updating the Git repository...'

        git add --all
        if ($LASTEXITCODE -ne 0) {
            Write-Error "Something went wrong updating the Git index with all changes."
        }

        $GitCommitDate = Get-Date -UFormat "%d/%m/%Y"
        git commit -m "Changes for $GitCommitDate"
        if ($LASTEXITCODE -ne 0) {
            Write-Error "Something went wrong committing all changes in the Git index."
        }

        git pull
        if ($LASTEXITCODE -ne 0) {
            Write-Error "Something went wrong pulling from the Git repository."
        }

        git push
        if ($LASTEXITCODE -ne 0) {
            Write-Error "Something went wrong pushing to the Git repository."
        }
    }

    # Although technically we don't need x64 we have only tested on it
    Test-Windows64bit

    # Check Git is installed on the system and setup the environment
    Test-GitInstalled

    # Check we're currently operating within a Git repository
    Test-GitRepository

    # Commit all changes and update the Git repository
    Update-GitRepository
}
