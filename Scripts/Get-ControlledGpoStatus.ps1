#Requires -Version 3.0

[CmdletBinding()]
Param()

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

$Results = @()
$TypeName = 'PSWinGlue.ControlledGpoStatus'

Update-TypeData -TypeName $TypeName -DefaultDisplayPropertySet @('Name', 'Status') -Force

# Retrieve domain GPOs and AGPM controlled GPOs
try {
    $DomainGPOs = Get-GPO -All
    $AgpmGPOs = Get-ControlledGpo
} catch {
    throw $_
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
    if ([Int32]$AgpmGPO.ComputerVersion -lt $DomainGPO.Computer.DSVersion) {
        $Result.Status += @('Domain computer policy is newer (Import)')
    } elseif ([Int32]$AgpmGPO.ComputerVersion -gt $DomainGPO.Computer.DSVersion) {
        $Result.Status += @('AGPM computer policy is newer (Deploy)')
    }

    # Check user policy is in sync
    #
    # The casting is necessary as the AGPM version properties are strings.
    if ([Int32]$AgpmGPO.UserVersion -lt $DomainGPO.User.DSVersion) {
        $Result.Status += @('Domain user policy is newer (Import)')
    } elseif ([Int32]$AgpmGPO.UserVersion -gt $DomainGPO.User.DSVersion) {
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
$MissingGPOs = $DomainGPOs | Where-Object { $_.Id -notin $AgpmGPOs.ID.TrimStart('{').TrimEnd('}') }
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