@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ============================================
echo    STT NOVA-HR Hub  -  Update System (one-click)
echo ============================================
echo.

echo [1/4] Sync backend code -^> deploy ...
copy /Y "src\jobtrack_apps_script.gs" "deploy\jobtrack_apps_script.gs" >nul
copy /Y "src\appsscript.json" "deploy\appsscript.json" >nul
copy /Y "nova-hr-hub\Recheck.gs" "deploy\Recheck.gs" >nul
echo.

echo [2/4] Backup to GitHub (skip if no git) ...
git rev-parse --is-inside-work-tree >nul 2>&1
if errorlevel 1 (
  echo    [skip] this folder is not a git repo yet
) else (
  git add -A
  git commit -m "update %date% %time%"
  git push
)
echo.

echo [3/4] Upload code to Apps Script (clasp push) ...
cd deploy
call clasp push -f
if errorlevel 1 goto PUSHFAIL
echo.

echo [4/4] Publish new version (clasp deploy) ...
call clasp deploy -i AKfycbyqyUP-2oE2PM_AzihyIq0_dMuZol4_lmTeRGFGeZ_ZMfc2gJTcje5rActFXW_OggfE -d "auto update"
if errorlevel 1 goto DEPLOYFAIL
cd ..
echo.

echo ============================================
echo    DONE!  Code uploaded + new version published.
echo    URL /exec is the SAME (link never changes).
echo    NEXT: close app tabs, open fresh, Ctrl+Shift+R
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
echo *** ERROR: clasp deploy FAILED - live app was NOT updated.
echo *** Manual fix: Apps Script - Deploy - Manage deployments - Edit - New version - Deploy
echo.
pause
exit /b 1
