@echo off
cd /d "%~dp0"
echo ============================================
echo    STT NOVA-HR Hub  -  Update (one-click)
echo ============================================
echo.

echo [1/4] Backup to GitHub ...
git add -A
git commit -m "update hub %date% %time%"
git push
echo.

echo [2/4] Push Hub code to Apps Script (clasp) ...
cd nova-hr-hub
call clasp push -f
if errorlevel 1 goto PUSHFAIL
echo.

echo [3/4] Publish new version to the SAME web app URL ...
call clasp deploy -i AKfycbyo0KxsnisfS65YNWeoRsgP2_j7ogBcG0S2-rpsEuYz9c7cNLFdj-kr9XvI_-y1gAUHPg -d "auto update"
if errorlevel 1 goto DEPLOYFAIL
cd ..
echo.

echo ============================================
echo    DONE! Code uploaded + web app updated.
echo    URL /exec is the SAME (link never changes).
echo    Refresh the dashboard page (Ctrl+Shift+R).
echo ============================================
echo.
pause
exit /b 0

:PUSHFAIL
echo.
echo *** ERROR: clasp push FAILED - code was NOT uploaded.
echo *** Common cause = clasp login expired. Fix: type  clasp login
echo.
pause
exit /b 1

:DEPLOYFAIL
echo.
echo *** ERROR: clasp deploy FAILED - web app was NOT updated.
echo *** Manual: Apps Script - Deploy - Manage deployments - Edit - New version - Deploy
echo.
pause
exit /b 1
