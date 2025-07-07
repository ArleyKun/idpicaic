@echo off
title id pic gen by yel
:loop
cls
cd /d "%~dp0"
py idgen.py
echo.
echo Script finished. Restarting...
timeout /t 2 >nul
goto loop
