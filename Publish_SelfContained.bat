@echo off
echo ========================================
echo Building Self-Contained ID Monitor
echo (No .NET Runtime Required!)
echo ========================================
echo.

cd /d "%~dp0"

echo [1/3] Cleaning previous build...
if exist "publish" rmdir /s /q "publish"
dotnet clean -c Release

echo.
echo [2/3] Publishing self-contained app...
echo This may take 1-2 minutes...
echo.

dotnet publish -c Release -r win-x64 --self-contained true -o publish

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ❌ BUILD FAILED!
    echo Please check for errors above.
    pause
    exit /b 1
)

echo.
echo [3/3] Build complete!
echo.
echo ========================================
echo Output files are in: publish\
echo ========================================
echo.
echo Main file: AutodeskIDMonitor.exe (~70-150 MB)
echo.
echo Copy ALL files from 'publish' folder to InnoSetup folder
echo Then run Build_Installer.bat
echo.
pause
