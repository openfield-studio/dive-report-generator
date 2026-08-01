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
    echo Pythonが見つかりません。
    echo https://www.python.org/downloads/ からインストールしてください（インストール時に「Add python.exe to PATH」にチェック）。
    pause
    exit /b 1
)

"%PY%" -c "import win32com.client, jpholiday" >nul 2>nul
if not !ERRORLEVEL! == 0 (
    echo 初回起動のため必要なライブラリをインストールします（pywin32, jpholiday）...
    "%PY%" -m pip install --quiet pywin32 jpholiday
    if not !ERRORLEVEL! == 0 (
        echo ライブラリのインストールに失敗しました。手動で以下を実行してください:
        echo   pip install pywin32 jpholiday
        pause
        exit /b 1
    )
)

start "" "%PY%" dive_record_generator.py
