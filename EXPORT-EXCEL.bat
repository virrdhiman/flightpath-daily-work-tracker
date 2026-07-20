@echo off
setlocal
cd /d "%~dp0"

echo ============================================================
echo   FlightPath Daily Work — Excel Export
echo ============================================================
echo.
echo   1^) Full log (8 Apr - 20 Jul 2026)
echo   2^) One month  (e.g. 2026-07)
echo   3^) One date   (e.g. 2026-07-20)
echo.

set /p CHOICE=Choose 1 / 2 / 3: 

if "%CHOICE%"=="2" (
  set /p MONTH=Month YYYY-MM: 
  python export_excel.py --month %MONTH%
  goto DONE
)
if "%CHOICE%"=="3" (
  set /p DAY=Date YYYY-MM-DD: 
  python export_excel.py --date %DAY%
  goto DONE
)

python export_excel.py

:DONE
echo.
echo Files are in the exports\ folder.
pause
