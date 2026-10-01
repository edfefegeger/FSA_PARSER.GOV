@echo off
chcp 65001 >nul
cd /d "%~dp0"
python fsa_parser.py %*
echo.
pause
