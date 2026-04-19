@echo off
REM Windows launcher for statsbook-print.py
REM Double-click this file after clicking "Update" in CRG scoreboard.
cd /d "%~dp0"
python statsbook-print.py
pause
