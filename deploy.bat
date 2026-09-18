@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo.
echo ============================================
echo  1. What changed
echo ============================================
git status --short

echo.
echo ============================================
echo  2. Commit
echo ============================================
git add -A
git commit -m "update: latest financial model build"

echo.
echo ============================================
echo  3. Push to GitHub
echo ============================================
git push

echo.
echo ============================================
echo  DONE.
echo.
echo  Wait 1-2 minutes, then open in a private
echo  window and check the notice bar at the top:
echo.
echo    https://dongjakjumin.github.io/baebaepro/
echo ============================================
echo.
pause
