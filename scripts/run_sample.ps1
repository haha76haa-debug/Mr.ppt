# 一键示例：生成示例 deck，并用 PowerPoint 真机导出前 3 页（Windows + Office）
#
#   powershell -NoProfile -ExecutionPolicy Bypass -File scripts/run_sample.ps1
#
$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
Set-Location $root

New-Item -ItemType Directory -Force build | Out-Null

Write-Host "[1/2] 生成示例 deck ..."
$env:PYTHONPATH = Join-Path $root "src"
python -m pptx_native sample build/sample.pptx

Write-Host "[2/2] PowerPoint 真机导出 1-3 页 ..."
foreach ($i in 1, 2, 3) {
    powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot "export_slide.ps1") `
        -Path (Join-Path $root "build/sample.pptx") `
        -Out (Join-Path $root "build/sample-p$i.png") -Slide $i
}

Write-Host "完成：build/sample.pptx / build/sample-p1..3.png"
