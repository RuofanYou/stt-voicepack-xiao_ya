@echo off
setlocal EnableDelayedExpansion
chcp 65001 >nul
title STT 可爱少女 语音包 卸载

set "HERE=%~dp0"
set "DLL=%HERE%stt_xiao_ya_sapi5.dll"
set "LOGFILE=%HERE%uninstall.log"

echo. > "%LOGFILE%"
echo 卸载 STT 可爱少女 语音包
echo.

net session >nul 2>&1
if errorlevel 1 (
    echo 需要管理员权限，请右键 -- 以管理员身份运行 --
    pause
    exit /b 1
)

if exist "%DLL%" (
    regsvr32 /u /s "%DLL%" >> "%LOGFILE%" 2>&1
    echo [OK] 已注销 COM DLL
) else (
    echo [警告] DLL 不存在，跳过 regsvr32 注销
)

echo.
echo 已清理注册表项。你可以手动删除这个文件夹以完全移除。
echo.
pause
