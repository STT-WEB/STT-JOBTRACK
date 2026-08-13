@echo off
cd /d "%~dp0"
echo ============================================
echo    STT JOBCOST  -  Pull monthly data (one-click)
echo ============================================
echo.

echo [1/2] Push latest code to Apps Script ...
cd nova-hr-hub
call clasp push -f
if errorlevel 1 goto PUSHFAIL
echo.

echo [2/2] Run rebuildJobcostFromMonthly ...
call clasp run rebuildJobcostFromMonthly
if errorlevel 1 goto RUNFAIL
cd ..
echo.

echo ============================================
echo    DONE! Data loaded into JOBCOST 2026.
echo    Open the Dashboard and press Refresh.
echo ============================================
echo.
pause
exit /b 0

:PUSHFAIL
echo.
echo *** ERROR: clasp push FAILED (clasp login expired?  type: clasp login)
echo.
pause
exit /b 1

:RUNFAIL
echo.
echo *** clasp run is not set up yet. Run it manually THIS ONCE:
echo ***  1) Open the Hub Apps Script project
echo ***  2) Pick function  rebuildJobcostFromMonthly
echo ***  3) Press Run
echo *** (One-time setup for one-click: turn ON the Apps Script API at
echo ***  https://script.google.com/home/usersettings  then this .bat works.)
echo.
pause
exit /b 1
