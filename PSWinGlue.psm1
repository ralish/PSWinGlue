# Import all functions
$Functions = Get-ChildItem (Join-Path $PSScriptRoot 'Functions') -File
foreach ($Function in $Functions) {
    . $Function.FullName
}
