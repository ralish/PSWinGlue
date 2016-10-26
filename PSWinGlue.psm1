# Import all functions
$Functions = Get-ChildItem -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Functions') -File
foreach ($Function in $Functions) {
    . $Function.FullName
}
