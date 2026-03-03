$ErrorActionPreference = "Stop"
$root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
Set-Location $root

$targets = @(
    ".\\build",
    ".\\dist\\KPI_Tool",
    ".\\logs",
    ".\\output",
    ".\\outputs",
    ".\\历史数据",
    ".\\输出结果"
)

foreach ($t in $targets) {
    if (Test-Path $t) {
        Write-Host "[CLEAN] $t"
        Remove-Item -Recurse -Force $t
    }
}

Write-Host "[DONE] clean outputs finished."

