@echo off
setlocal EnableExtensions
set "QUARTAL_HOST=0.0.0.0"
call "%~dp0run_windows.bat"
exit /b %ERRORLEVEL%
