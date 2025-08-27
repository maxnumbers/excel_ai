#!/bin/bash
echo "Building Local AI Excel Add-in..."
echo

# Check if .NET SDK is available
if ! command -v dotnet &> /dev/null; then
    echo "ERROR: .NET SDK not found"
    echo "Please install .NET Framework 4.7.2 or later SDK"
    exit 1
fi

echo "[1/3] Restoring NuGet packages..."
cd LocalAI.Excel.AddIn
dotnet restore

echo
echo "[2/3] Building Debug version..."
dotnet build -c Debug

echo
echo "[3/3] Building Release version (for distribution)..."
dotnet build -c Release

echo
echo "Build complete!"
echo
echo "Files created:"
echo "- Debug: LocalAI.Excel.AddIn/bin/Debug/LocalAI.Excel.AddIn-packed.xll"
echo "- Release: LocalAI.Excel.AddIn/bin/Release/LocalAI.Excel.AddIn-packed.xll"
echo
echo "To install:"
echo "1. Copy the .xll file to your desired location"
echo "2. In Excel: File > Options > Add-ins > Manage Excel Add-ins > Browse"
echo "3. Select the .xll file and click OK"
echo