@echo off

set TASK_NAME=Outlook Calendar Sync

schtasks /delete /tn "%TASK_NAME%" /f >nul 2>nul

schtasks /create ^
 /tn "%TASK_NAME%" ^
 /sc minute ^
 /mo 15 ^
 /tr "\"%~dp0update.bat\"" ^
 /f

echo Task created successfully.
pause
