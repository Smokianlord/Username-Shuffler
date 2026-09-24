@echo off
setlocal
echo ===================================================
echo  Username Shuffler - Windows Standalone Build Script
echo ===================================================
python build.py
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERROR] Build failed! Check the error log above.
    pause
    exit /b %ERRORLEVEL%
)
echo.
echo [SUCCESS] Build completed! Check the 'dist' folder.
pause
