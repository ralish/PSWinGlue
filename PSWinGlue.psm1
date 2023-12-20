# See the help for Set-StrictMode for the full details on what this enables.
Set-StrictMode -Version 2.0

# Import all scripts as functions
$Scripts = Get-ChildItem -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Scripts') -File
foreach ($Script in $Scripts) {
    $FunctionName = $Script.BaseName
    $FunctionPath = 'Function:\{0}' -f $FunctionName

    if (Test-Path -Path $FunctionPath) {
        Write-Warning -Message ('Skipping import of existing function: {0}' -f $FunctionName)
        continue
    }

    New-Item -Path $FunctionPath -Value (Get-Content -Path $Script.FullName -Raw)
}
