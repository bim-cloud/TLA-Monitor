@echo off
echo =====================================================
echo  Tangent ID Monitor - Clean Build
echo =====================================================
echo.

echo Cleaning old build files...
if exist "bin" rmdir /s /q bin
if exist "obj" rmdir /s /q obj
echo Done.
echo.

echo Restoring packages...
dotnet restore
echo.

echo Building project...
dotnet build -c Release
echo.

if %ERRORLEVEL% EQU 0 (
    echo =====================================================
    echo  BUILD SUCCESSFUL!
    echo  Designer errors should now be resolved.
    echo  If errors persist in Visual Studio, close and reopen
    echo  the solution.
    echo =====================================================
) else (
    echo =====================================================
    echo  BUILD FAILED - Check errors above
    echo =====================================================
)
echo.
pause
