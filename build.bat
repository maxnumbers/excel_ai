@echo off
echo Building Local AI Excel Add-in...
echo.

REM Check if .NET SDK is available
dotnet --version >nul 2>&1
if %errorLevel% neq 0 (
    echo ERROR: .NET SDK not found
    echo Please install .NET Framework 4.7.2 or later SDK
    pause
    exit /b 1
)

echo [1/3] Restoring NuGet packages...
cd LocalAI.Excel.AddIn
dotnet restore

echo.
echo [2/3] Building Debug version...
dotnet build -c Debug

echo.
echo [3/3] Building Release version (for distribution)...
dotnet build -c Release

echo.
echo Build complete!
echo.
echo Files created:
echo - Debug: LocalAI.Excel.AddIn\bin\Debug\LocalAI.Excel.AddIn-packed.xll
echo - Release: LocalAI.Excel.AddIn\bin\Release\LocalAI.Excel.AddIn-packed.xll
echo.
echo To install:
echo 1. Copy the .xll file to your desired location
echo 2. In Excel: File > Options > Add-ins > Manage Excel Add-ins > Browse
echo 3. Select the .xll file and click OK
echo.
pause