@echo off
title Launch Microsoft Teams + Next Meeting Tile
set "APP_DIR=%~dp0"
set "EXE_PATH=%APP_DIR%bin\Release\net9.0-windows\TeamsNextMeetingTile.exe"
set "PORTABLE_EXE=%APP_DIR%Portable\TeamsNextMeetingTile.exe"
set "LOCAL_EXE=%APP_DIR%TeamsNextMeetingTile.exe"

:: 1. Start Microsoft Teams (via protocol handler so it works with New Teams MSIX)
start "" "msteams:"

:: 2. Start Teams Next Meeting Tile
if exist "%LOCAL_EXE%" (
    start "" "%LOCAL_EXE%"
    exit /b 0
)

if exist "%EXE_PATH%" (
    start "" "%EXE_PATH%"
    exit /b 0
)

if exist "%PORTABLE_EXE%" (
    start "" "%PORTABLE_EXE%"
    exit /b 0
)
