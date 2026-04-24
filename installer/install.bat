@echo off
setlocal EnableDelayedExpansion
chcp 65001 >nul
title STT 可爱少女 语音包 安装

set "HERE=%~dp0"
set "DLL=%HERE%stt_xiao_ya_sapi5.dll"
set "LOGFILE=%HERE%install.log"

echo. > "%LOGFILE%"
echo ======================================================== | tee_fallback
echo   STT 可爱少女 (xiao_ya) 中文语音包 - 安装
echo ========================================================
echo.
echo 日志文件: %LOGFILE%
echo.

echo [1/4] 检查管理员权限...
net session >nul 2>&1
if errorlevel 1 (
    echo [X] 当前不是管理员权限，安装系统语音必须用管理员运行。
    echo     请右键此 install.bat -- 以管理员身份运行 --
    echo.
    pause
    exit /b 1
)
echo [OK] 已获得管理员权限

echo [2/4] 检查文件完整性...
if not exist "%DLL%" (
    echo [X] 找不到 stt_xiao_ya_sapi5.dll，包可能不完整。
    pause
    exit /b 1
)
if not exist "%HERE%xiao_ya\zh_CN-xiao_ya-medium.onnx" (
    echo [X] 找不到 xiao_ya\zh_CN-xiao_ya-medium.onnx 模型文件。
    pause
    exit /b 1
)
echo [OK] 文件完整

echo [3/4] 注册 SAPI5 COM DLL 到系统...
regsvr32 /s "%DLL%" >> "%LOGFILE%" 2>&1
if errorlevel 1 (
    echo [X] regsvr32 注册失败 (exit %errorlevel%^)。见 %LOGFILE%
    pause
    exit /b 1
)
echo [OK] COM DLL 已注册

echo [4/4] 验证 SAPI5 语音列表...
powershell -NoProfile -Command ^
    "Add-Type -AssemblyName System.Speech;" ^
    "$s = New-Object System.Speech.Synthesis.SpeechSynthesizer;" ^
    "$voices = $s.GetInstalledVoices() ^| Where-Object { $_.VoiceInfo.Name -match 'xiao_ya' -or $_.VoiceInfo.Name -match '可爱少女' };" ^
    "if ($voices) { Write-Host '[OK] 已在系统语音列表中找到：'; $voices ^| ForEach-Object { Write-Host '    -' $_.VoiceInfo.Name } } else { Write-Host '[警告] 未在 SAPI5 枚举中找到，但注册表已写入。WoW 可能仍能看到。' }"

echo.
echo ========================================================
echo   安装完成
echo ========================================================
echo.
echo 下一步：
echo   1. 重启 WoW
echo   2. 打开 STT 设置 (/st) 的 TTS 语音选择
echo   3. 选择 "STT 可爱少女 (xiao_ya)"
echo.
echo 卸载请运行同目录的 uninstall.bat
echo.
pause
