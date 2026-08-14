@echo off
cd /d "%~dp0"
echo ============================================
echo    Login clasp (Apps Script) - fix invalid_rapt
echo ============================================
echo.
echo A browser window will open.
echo Login with the SAME Google account used for
echo Apps Script (Hub / JOBTRACK), then click Allow.
echo.
call clasp login
echo.
echo ============================================
echo    If you saw "Authorization successful" =
echo    DONE. Now double-click  อัปเดต-Hub.bat  again.
echo ============================================
echo.
pause
