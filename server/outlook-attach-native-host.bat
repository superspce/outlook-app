@echo off
REM Wrapper script for Windows Native Messaging Host
REM This is needed because Chrome Native Messaging doesn't support command-line arguments

REM Get the directory where this batch file is located
set SCRIPT_DIR=%~dp0

REM Find Python interpreter
for /f "delims=" %%i in ('where python 2^>nul') do set PYTHON_PATH=%%i

if "%PYTHON_PATH%"=="" (
    echo Error: Python not found in PATH >&2
    exit /b 1
)

REM Run the Python script with stdin/stdout connected
"%PYTHON_PATH%" "%SCRIPT_DIR%outlook-attach-native-host.py"

