@echo off
REM One-click build: convert input MD into a corporate Word doc.
REM Usage: run.bat [config.json]
REM Prefers the project virtualenv (.venv); falls back to system python.
setlocal
set "CONFIG=%~1"
if "%CONFIG%"=="" set "CONFIG=config\md2docx.json"
cd /d "%~dp0"

set "PY=python"
if exist ".venv\Scripts\python.exe" set "PY=.venv\Scripts\python.exe"

"%PY%" src\md2docx.py --config "%CONFIG%"
endlocal
