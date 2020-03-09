$ConsoleAPI = Get-Content -Raw -Path (Join-Path -Path $PSScriptRoot -ChildPath 'ConsoleAPI.cs')

if (!('PSWinGlue.ConsoleAPI' -as [Type])) {
    Add-Type -Namespace PSWinGlue -Name ConsoleAPI -MemberDefinition $ConsoleAPI
} else {
    Write-Warning -Message 'Unable to add ConsoleAPI type as it already exists.'
}
