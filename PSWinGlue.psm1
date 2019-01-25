# Import all scripts as functions
$Scripts = Get-ChildItem -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Scripts') -File
foreach ($Script in $Scripts) {
    New-Item -Path ('Function:\{0}' -f $Script.BaseName) -Value (Get-Content -Path $Script.FullName -Raw)
}
