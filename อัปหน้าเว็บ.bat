@echo off
cd /d "%~dp0"
echo ============================================
echo    STT JOBTRACK  -  Update Web (one-click)
echo ============================================
echo.

echo [1/3] Sync latest from GitHub (pull) ...
git pull --no-edit
if errorlevel 1 goto PULLFAIL
echo.

echo [2/3] Copy web file (src -^> repo root) ...
copy /Y "src\job_checkin_app.html" "job_checkin_app.html" >nul
if errorlevel 1 goto COPYFAIL
echo.

echo [3/3] Push to GitHub Pages ...
git add -A
git commit -m "update web %date% %time%"
git push
if errorlevel 1 goto PUSHFAIL
echo.

echo ============================================
echo    DONE! Web pushed to GitHub.
echo    Wait 1-2 min, then open the app on the
echo    phone and hard-refresh (clear cache).
echo    URL is the SAME (link never changes).
echo ============================================
echo.
pause
exit /b 0

:PULLFAIL
echo.
echo *** ERROR: git pull FAILED (login/internet, or a conflict).
echo *** Send this screen to Candy before continuing.
echo.
pause
exit /b 1

:COPYFAIL
echo.
echo *** ERROR: cannot find  src\job_checkin_app.html
echo.
pause
exit /b 1

:PUSHFAIL
echo.
echo *** ERROR: git push FAILED (GitHub login or internet?).
echo *** You can still upload job_checkin_app.html manually on GitHub.
echo.
pause
exit /b 1
