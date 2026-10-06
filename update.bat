@echo off
REM Update the Outlook Calendar and commit changes to Git

REM Change the current directory to the location of this script
CD /d "%~dp0"

REM Export Outlook Calendar to calendar.ics
python calendar_export.py

REM Stage the updated calendar.ics file for commit
git add calendar.ics

REM Check if there are any staged changes
git diff --cached --quiet

REM If there are no staged changes, skip the commit
if %errorlevel%==0 goto END_SCRIPT

REM Commit and push the changes if there are any
git commit -m "Update calendar.ics"
git push

:END_SCRIPT
echo Error level: %errorlevel%
echo No update to commit