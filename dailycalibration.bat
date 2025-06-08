@echo off
cd C:\Calibrate
powerShell -File calibrate.ps1 >> "calibrationlog.txt"
exit /b 0
