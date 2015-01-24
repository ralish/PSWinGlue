# Ensure all modules files are unblocked
Get-ChildItem -Path $PSScriptRoot | Unblock-File
# Source in all the module script files
Get-ChildItem -Path $PSScriptRoot\*.ps1 | ForEach-Object { . $_.FullName }
