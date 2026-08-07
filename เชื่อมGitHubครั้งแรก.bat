@echo off
cd /d "%~dp0"
echo ============================================
echo   Connect / finish sync with GitHub
echo   SAFE: employee web files are NOT changed
echo ============================================
echo.

echo [1/8] Set git identity (this repo only) ...
git config user.email >nul 2>&1
if errorlevel 1 git config user.email "sasipa@suteetankers.com"
git config user.name >nul 2>&1
if errorlevel 1 git config user.name "BEER"

echo [2/8] Ensure local git ...
if not exist ".git" git init

echo [3/8] Ensure GitHub link ...
git remote get-url origin >nul 2>&1
if errorlevel 1 git remote add origin https://github.com/stt-web/STT-JOBTRACK.git

echo [4/8] Fetch live files from GitHub ...
git fetch origin
if errorlevel 1 goto FETCHFAIL

echo [5/8] Switch to main branch ...
git checkout main 2>nul
if errorlevel 1 git checkout -b main --track origin/main

echo [6/8] Stage your project files ...
git add -A

echo [7/8] Save (commit) ...
git commit -m "Add full project backup (web files unchanged)"

echo [8/8] Upload to GitHub ...
git push -u origin main
if errorlevel 1 goto PUSHFAIL

echo.
echo ============================================
echo   DONE! PC and GitHub are now in sync.
echo   Employee app is untouched (same link).
echo ============================================
echo.
pause
exit /b 0

:FETCHFAIL
echo.
echo *** ERROR: could not reach GitHub (fetch failed).
echo *** Check internet, then run this file again.
echo.
pause
exit /b 1

:PUSHFAIL
echo.
echo *** Saved locally, but upload to GitHub FAILED.
echo *** Usually = GitHub login needed (a login window may pop up).
echo *** Log in to GitHub, then run this file again to finish.
echo.
pause
exit /b 1
