param(
    [Parameter(Mandatory = $false)]
    [string]$Version = "v0.1.0"
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
Set-Location $root

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$zipName = "Event-Support-Tool-$Version-win64.zip"
$bundleDir = Join-Path $root "dist\\release_$($Version)_$timestamp"

Write-Host "[STEP] install dependencies"
python -m pip install --upgrade pip
pip install -r .\requirements.txt
pip install pyinstaller

Write-Host "[STEP] build exe"
pyinstaller .\scripts\packaging\KPI_Tool.spec --noconfirm --clean --distpath .\dist --workpath .\build

if (!(Test-Path ".\\dist\\KPI_Tool")) {
    throw "Build failed: .\\dist\\KPI_Tool not found"
}

Write-Host "[STEP] prepare release bundle"
New-Item -ItemType Directory -Force -Path $bundleDir | Out-Null
Copy-Item -Recurse -Force .\dist\KPI_Tool\* $bundleDir

# 附带示例配置和模板（非敏感）
Copy-Item -Recurse -Force .\configs (Join-Path $bundleDir "configs")
Copy-Item -Recurse -Force .\assets\templates (Join-Path $bundleDir "assets_templates")

@'
快速使用说明
1) 双击 KPI_Tool.exe 运行
2) 按需准备输入目录与配置
3) 输出结果请勿直接提交到代码仓库
'@ | Set-Content -Encoding UTF8 (Join-Path $bundleDir "快速使用说明.txt")

Write-Host "[STEP] zip package"
if (Test-Path ".\\dist\\$zipName") { Remove-Item ".\\dist\\$zipName" -Force }
Compress-Archive -Path "$bundleDir\\*" -DestinationPath ".\\dist\\$zipName"

Write-Host "[DONE] release package: dist\\$zipName"

