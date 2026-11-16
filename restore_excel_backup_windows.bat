@echo off
setlocal EnableExtensions
chcp 65001 >nul
cd /d "%~dp0"
echo ============================================================
echo  KVARTAL - RESTORE LAST GOOD EXCEL FILE
echo ============================================================
echo.
if not exist "data\active.lastgood.xlsx" (
  echo ERROR: data\active.lastgood.xlsx was not found.
  echo Copy active.lastgood.xlsx from your previous project folder
  echo into this project's data folder, then run this file again.
  echo.
  pause
  exit /b 1
)
if exist "data\active.xlsx" (
  copy /y "data\active.xlsx" "data\active.before_restore.xlsx" >nul
)
copy /y "data\active.lastgood.xlsx" "data\active.xlsx" >nul
if errorlevel 1 (
  echo ERROR: Could not restore the backup file.
  pause
  exit /b 1
)
echo Restored data\active.xlsx from data\active.lastgood.xlsx.
echo A previous active file, if present, was saved as active.before_restore.xlsx.
echo.
pause
