@echo off
:loop
cd /d "%~dp0"
py word.py
echo.
echo Script finished. Restarting...
timeout /t 2 >nul
goto loop
