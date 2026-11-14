# 用 PowerPoint 本身把某一页导出成 PNG —— 唯一可信的“真机渲染”。
#
#   powershell -NoProfile -ExecutionPolicy Bypass -File export_slide.ps1 `
#       -Path "C:\path\deck.pptx" -Out "C:\tmp\slide6.png" -Slide 6 -Width 2001 -Height 1126
#
# 说明：
#   * 必须 Visible = msoTrue，否则超大 deck 常报 “PowerPoint could not open the file”
#   * 打开用 WithWindow=$true；ReadOnly=$true 防止改动源文件
#   * 结束务必 Quit，避免后台残留 POWERPNT 进程占住文件
param(
    [Parameter(Mandatory = $true)][string]$Path,
    [Parameter(Mandatory = $true)][string]$Out,
    [int]$Slide = 1,
    [int]$Width = 2001,
    [int]$Height = 1126
)
$ErrorActionPreference = "Stop"
$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
try {
    Start-Sleep -Milliseconds 300
    $pres = $ppt.Presentations.Open($Path, $true, $false, $true)
    "opened: $($pres.Slides.Count) slides"
    $pres.Slides.Item($Slide).Export($Out, "PNG", $Width, $Height)
    "exported -> $Out"
    $pres.Close()
}
catch {
    "FAILED: $($_.Exception.Message)"
    exit 1
}
finally {
    try { $ppt.Quit() } catch {}
    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null } catch {}
}
