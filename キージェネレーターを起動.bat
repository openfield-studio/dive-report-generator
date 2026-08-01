@echo off
setlocal enabledelayedexpansion
cd /d "%~dp0"

set "PY="
for %%P in (pythonw.exe python.exe) do (
    if not defined PY (
        where %%P >nul 2>nul && set "PY=%%P"
    )
)
if not defined PY (
    where py >nul 2>nul
    if !ERRORLEVEL! == 0 set "PY=py"
)

if not defined PY (
    echo Python‚ªŒ©‚Â‚©‚è‚Ü‚¹‚ñB
    pause
    exit /b 1
)

start "" "%PY%" key_generator.py
