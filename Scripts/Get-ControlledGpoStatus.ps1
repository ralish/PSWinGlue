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
        Status      = 'Unknown'
    }

    $DomainGPO = $DomainGPOs | Where-Object { $_.DisplayName -eq $AgpmGPO.Name }
    if ($DomainGPO) {
        $Result.Domain = $DomainGPO

        # The casting is necessary as the AGPM version properties are strings
        if ([Int32]$AgpmGPO.ComputerVersion -eq $DomainGPO.Computer.DSVersion -and [Int32]$AgpmGPO.UserVersion -eq $DomainGPO.User.DSVersion) {
            $Result.Status = 'Current'
        } elseif ([Int32]$AgpmGPO.ComputerVersion -le $DomainGPO.Computer.DSVersion -and [Int32]$AgpmGPO.UserVersion -le $DomainGPO.User.DSVersion) {
            $Result.Status = 'Out-of-date (Import)'
        } elseif ([Int32]$AgpmGPO.ComputerVersion -ge $DomainGPO.Computer.DSVersion -and [Int32]$AgpmGPO.UserVersion -ge $DomainGPO.User.DSVersion) {
            $Result.Status = 'Newer (Deploy)'
        } else {
            $Result.Status = 'Inconsistent'
        }
    } else {
        $Result.Status = 'Only exists in AGPM'
    }

    $Results += $Result
}

# Check the status of any domain GPOs not controlled by AGPM
$MissingGPOs = $DomainGPOs | Where-Object { $_.DisplayName -notin $AgpmGPOs.Name }
foreach ($MissingGPO in $MissingGPOs) {
    $Results += [PSCustomObject]@{
        PSTypeName  = $TypeName
        Name        = $MissingGPO.DisplayName
        AGPM        = $null
        Domain      = $MissingGPO
        Status      = 'Only exists in Domain'
    }
}

return ($Results | Sort-Object -Property Name)
