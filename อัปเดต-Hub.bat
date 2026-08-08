@echo off
cd /d "%~dp0"
echo ============================================
echo    STT NOVA-HR Hub  -  Update code (one-click)
echo ============================================
echo.

echo [1/3] Backup to GitHub ...
git add -A
git commit -m "update hub %date% %time%"
git push
echo.

echo [2/3] Push Hub code to Apps Script (clasp) ...
cd nova-hr-hub
call clasp push -f
if errorlevel 1 goto PUSHFAIL
cd ..
echo.

echo ============================================
echo    DONE! Hub code uploaded.
echo    Open the Hub project and run the function you need
echo    (recheckJanuary / computeJobcostMonth ...).
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
