param(
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$ArgsList
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
Set-Location $root

Write-Host "[RUN] python .\\src\\backend_logic.py $($ArgsList -join ' ')"
python .\src\backend_logic.py @ArgsList

