#Requires -Version 3.0

[CmdletBinding()]
Param(
    [ValidateNotNullOrEmpty()]
    [String]$Domain,

    [ValidateNotNullOrEmpty()]
    [String]$AgpmServer
)

# Check required modules are present:
# - GroupPolicy: Provided by Windows Server GPMC feature or Windows Client RSAT feature
# - Microsoft.Agpm: Included with AGPM Client
$RequiredModules = @('GroupPolicy', 'Microsoft.Agpm')
foreach ($Module in $RequiredModules) {
    Write-Verbose -Message ('Checking module is available: {0}' -f $Module)
    if (!(Get-Module -Name $Module -ListAvailable)) {
        throw ('Required module not available: {0}' -f $Module)
    }
}

if ($Domain -and !$AgpmServer) {
    $AgpmServer = 'agpm.{0}' -f $Domain
    Write-Warning -Message ('Using default AGPM server: {0}' -f $AgpmServer)
}

$Results = @()
$TypeName = 'PSWinGlue.ControlledGpoStatus'

Update-TypeData -TypeName $TypeName -DefaultDisplayPropertySet @('Name', 'Status') -Force

# Retrieve domain GPOs
try {
    if ($Domain) {
        $DomainGPOs = @(Get-GPO -All -Domain $Domain)
    } else {
        $DomainGPOs = @(Get-GPO -All)
    }
} catch {
    throw $_
}

# Retrieve AGPM GPOs
try {
    $DefaultArchiveChanged = $false

    # Overriding the DefaultArchive registry value is unfortunately necessary
    # as the Microsoft.Agpm PowerShell module does not respect any configured
    # per-domain servers listed under the Domains key, unlike the MMC snap-in.
    if ($AgpmServer) {
        $AgpmRegPath = 'HKLM:\Software\Microsoft\AGPM'

        if (Test-Path -Path $AgpmRegPath) {
            $AgpmReg = Get-Item -Path $AgpmRegPath -ErrorAction Stop
        } else {
            $AgpmReg = New-Item -Path $AgpmRegPath -ErrorAction Stop
        }

        if ($AgpmReg.Property.Contains('DefaultArchive')) {
            $OriginalDefaultArchive = $AgpmReg.GetValue('DefaultArchive')
        }

        if ($AgpmServer -notmatch ':[1-9][0-9]*$') {
            $DefaultArchive = '{0}:4600' -f $AgpmServer
        } else {
            $DefaultArchive = $AgpmServer
        }

        Set-ItemProperty -Path $AgpmRegPath -Name 'DefaultArchive' -Value $DefaultArchive -ErrorAction Stop
        $DefaultArchiveChanged = $true
    }

    if ($Domain) {
        $AgpmGPOs = @(Get-ControlledGpo -Domain $Domain -ErrorAction Stop)
    } else {
        $AgpmGPOs = @(Get-ControlledGpo -ErrorAction Stop)
    }
} catch {
    throw $_
} finally {
    if ($DefaultArchiveChanged) {
        Set-ItemProperty -Path $AgpmRegPath -Name 'DefaultArchive' -Value $OriginalDefaultArchive
    }
}

# Check the status of all AGPM controlled GPOs
foreach ($AgpmGPO in $AgpmGPOs) {
    $Result = [PSCustomObject]@{
        PSTypeName  = $TypeName
        Name        = $AgpmGPO.Name
        AGPM        = $AgpmGPO
        Domain      = $null
        Status      = @()
    }

    $DomainGPO = $DomainGPOs | Where-Object { $_.Id -eq $AgpmGPO.ID.TrimStart('{').TrimEnd('}') }
    if ($DomainGPO) {
        $Result.Domain = $DomainGPO
    } else {
        $Result.Status = @('Only exists in AGPM')
        $Results += $Result
        continue
    }

    # Check display name is in sync
    if ($AgpmGPO.Name -ne $DomainGPO.DisplayName) {
        $Result.Status += @('Name mismatch')
    }

    # Check computer policy is in sync
    #
    # The casting is necessary as the AGPM version properties are strings.
    if ([Int]$AgpmGPO.ComputerVersion -lt $DomainGPO.Computer.DSVersion) {
        $Result.Status += @('Domain computer policy is newer (Import)')
    } elseif ([Int]$AgpmGPO.ComputerVersion -gt $DomainGPO.Computer.DSVersion) {
        $Result.Status += @('AGPM computer policy is newer (Deploy)')
    }

    # Check user policy is in sync
    #
    # The casting is necessary as the AGPM version properties are strings.
    if ([Int]$AgpmGPO.UserVersion -lt $DomainGPO.User.DSVersion) {
        $Result.Status += @('Domain user policy is newer (Import)')
    } elseif ([Int]$AgpmGPO.UserVersion -gt $DomainGPO.User.DSVersion) {
        $Result.Status += @('AGPM user policy is newer (Deploy)')
    }

    # Check WMI filter is in sync
    if ($AgpmGPO.WmiFilterName -or $DomainGPO.WmiFilter) {
        if ($AgpmGPO.WmiFilterName -ne $DomainGPO.WmiFilter.Name) {
            $Result.Status += @('WMI filter mismatch')
        }
    }

    if (!$Result.Status) {
        $Result.Status = @('OK')
    }

    $Results += $Result
}

# Add any domain GPOs not controlled by AGPM
if ($AgpmGPOs.Count -gt 0) {
    $MissingGPOs = $DomainGPOs | Where-Object { $_.Id -notin $AgpmGPOs.ID.TrimStart('{').TrimEnd('}') }
} else {
    $MissingGPOs = $DomainGPOs
}

foreach ($MissingGPO in $MissingGPOs) {
    $Results += [PSCustomObject]@{
        PSTypeName  = $TypeName
        Name        = $MissingGPO.DisplayName
        AGPM        = $null
        Domain      = $MissingGPO
        Status      = @('Only exists in Domain')
    }
}

return ($Results | Sort-Object -Property Name)
