@echo off
title 定时关闭所有程序
echo ===============================================
echo   定时关闭所有程序（兼容Windows 7 专业版 X32）
echo ===============================================
echo 注意：此操作将强制结束当前用户的所有进程
echo 未保存的数据将会丢失！请提前保存工作。
echo.
set /p minutes=请输入等待的分钟数（例如 5 表示5分钟后关闭）:

REM 检查输入是否为有效数字
echo %minutes%|findstr /r "^[0-9][0-9]*$" >nul
if errorlevel 1 (
    echo 输入无效，请输入一个正整数。
    pause
    exit /b
)

set /a seconds=minutes*60
echo 将在 %minutes% 分钟后关闭所有程序...
timeout /t %seconds% /nobreak >nul

echo 正在关闭所有程序...
taskkill /F /FI "USERNAME eq %USERNAME%"

echo 操作完成。
pause