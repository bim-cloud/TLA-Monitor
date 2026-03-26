@echo off
echo =====================================================
echo  Tangent ID Monitor - Fix Designer Errors
echo =====================================================
echo.

echo Step 1: Cleaning bin and obj folders...
if exist "bin" rmdir /s /q bin
if exist "obj" rmdir /s /q obj
echo Done.
echo.

echo Step 2: Restoring NuGet packages...
dotnet restore
echo Done.
echo.

echo Step 3: Building project...
dotnet build -c Debug
echo.

if %ERRORLEVEL% EQU 0 (
    echo =====================================================
    echo  BUILD SUCCESSFUL!
    echo  Please restart Visual Studio to clear designer cache.
    echo =====================================================
) else (
    echo =====================================================
    echo  BUILD FAILED - Check errors above
    echo =====================================================
)
echo.
pause
