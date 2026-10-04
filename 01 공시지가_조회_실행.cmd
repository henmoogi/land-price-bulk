@echo off
chcp 65001 >nul
rem Land price bulk lookup launcher. Helper = PowerShell script built into Windows (no Python needed).
rem API key: automation/vworld_key.txt. Set LAND_NO_OPEN=1 to start the helper only.
set "HELPER=%~dp0automation\공시지가_도우미.ps1"
set "HTML=%~dp0automation\공시지가_대량조회_배포.html"
if not exist "%HTML%" set "HTML=%~dp0공시지가_대량조회_배포.html"
powershell -NoProfile -ExecutionPolicy Bypass -Command "$ok=$false; try { Invoke-RestMethod -Uri 'http://127.0.0.1:43129/health' -TimeoutSec 1 | Out-Null; $ok=$true } catch {}; if(-not $ok){ $ps=(Get-Command pwsh.exe -ErrorAction SilentlyContinue).Source; if(-not $ps){ $ps=(Get-Command powershell.exe).Source }; Start-Process -FilePath $ps -ArgumentList @('-NoProfile','-ExecutionPolicy','Bypass','-WindowStyle','Hidden','-File',('\"' + $env:HELPER + '\"')) -WindowStyle Hidden; Start-Sleep -Seconds 3 }"
if /I "%LAND_NO_OPEN%"=="1" exit /b 0
start "" "%HTML%"
