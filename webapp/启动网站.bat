@echo off
chcp 65001 >nul
rem Windows 用户：双击本文件即可启动「通话复盘助手」网站。
cd /d "%~dp0.."
echo 正在启动通话复盘助手……
python webapp\server.py
echo.
echo 网站已停止，可以关闭本窗口。
pause
