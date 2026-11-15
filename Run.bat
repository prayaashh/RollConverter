@echo off
setlocal
cd /d "%~dp0"

rem Prefer the windowless Python launchers so no console is left behind.
where pyw.exe >nul 2>nul
if not errorlevel 1 (
    start "" pyw.exe -3 "rollconverter.py"
    exit /b 0
)

where pythonw.exe >nul 2>nul
if not errorlevel 1 (
    start "" pythonw.exe "rollconverter.py"
    exit /b 0
)

rem Fall back to the console interpreters.
where py.exe >nul 2>nul
if not errorlevel 1 (
    py.exe -3 "rollconverter.py"
    goto finished
)

where python.exe >nul 2>nul
if not errorlevel 1 (
    python.exe "rollconverter.py"
    goto finished
)

echo.
echo   Python 3 is required but was not found on this computer.
echo.
echo   1. Install it from https://www.python.org/downloads/
echo   2. Tick "Add python.exe to PATH" during setup
echo   3. Run this file again
echo.
pause
exit /b 1

:finished
if errorlevel 1 pause
exit /b 0
