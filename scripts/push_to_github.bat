@echo off
rem 提交并推送（Windows 双击即可）
setlocal
cd /d "%~dp0.."
echo ==== 当前状态 ====
git status --short
echo.
set /p msg="提交信息（直接回车用默认）: "
if "%msg%"=="" set msg=chore: update
git add -A
git commit -m "%msg%"
if errorlevel 1 echo （没有需要提交的改动）
echo.
echo ==== 推送 ====
git push
echo.
echo 完成（已推送到 origin）。
pause
