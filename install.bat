@echo off
echo ========================================
echo  Local AI Excel Add-in (Native) v2.0
echo  Simple Installation Script
echo ========================================
echo.

echo This installer will:
echo 1. Build the Excel-DNA add-in
echo 2. Guide you through Excel installation
echo 3. Help configure your AI service
echo.
pause

echo [1/3] Building the add-in...
if not exist "LocalAI.Excel.AddIn\LocalAI.Excel.AddIn.csproj" (
    echo ERROR: Project files not found
    echo Please run this from the extracted ZIP folder
    pause
    exit /b 1
)

call build.bat
if %errorLevel% neq 0 (
    echo Build failed! Check error messages above.
    pause
    exit /b 1
)

echo.
echo [2/3] Add-in built successfully!
echo.
echo Next steps:
echo 1. Open Excel
echo 2. File → Options → Add-ins → Manage Excel Add-ins → Go
echo 3. Browse and select:
echo    %CD%\LocalAI.Excel.AddIn\bin\Release\LocalAI.Excel.AddIn-packed.xll
echo 4. Click OK
echo.
echo [3/4] After loading in Excel:
echo 1. Look for "Local AI" tab in ribbon
echo 2. Click "AI Configuration" 
echo 3. Configure Ollama (localhost:11434) or LM Studio (localhost:1234)
echo 4. Test connection
echo.
echo [4/4] For function tooltips and parameter help:
echo 1. Download ExcelDna.IntelliSense64.xll from:
echo    https://github.com/Excel-DNA/IntelliSense/releases
echo 2. Load it in Excel the same way as this add-in
echo 3. Function tooltips will then appear automatically
echo.
echo Installation complete!
echo Try: =AI.TEST() first, then =AI.CHAT("Hello AI!")
echo.
echo Press any key to open the Release folder...
pause >nul
explorer "%CD%\LocalAI.Excel.AddIn\bin\Release"
