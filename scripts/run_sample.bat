@echo off
rem 生成示例 deck（Windows 双击即可）
setlocal
cd /d "%~dp0.."
set PYTHONPATH=%CD%\src
python -m pptx_native sample build\sample.pptx
if errorlevel 1 goto :fail
python -m pptx_native themes
echo.
echo 完成：build\sample.pptx
pause
exit /b 0
:fail
echo 生成失败，请确认已安装依赖（pip install -e .）
pause
exit /b 1
