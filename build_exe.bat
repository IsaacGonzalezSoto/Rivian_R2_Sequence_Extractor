@echo off
REM ============================================================================
REM R2 Sequence Extractor - Build Script
REM
REM This script builds a standalone Windows executable using PyInstaller.
REM The resulting .exe file will be located in the dist/ folder.
REM ============================================================================

echo ============================================================================
echo R2 SEQUENCE EXTRACTOR - BUILD EXECUTABLE
echo ============================================================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python is not installed or not in PATH
    echo Please install Python 3.7+ and try again
    pause
    exit /b 1
)

echo [1/4] Checking Python installation...
python --version
echo.

REM Check if PyInstaller is installed
echo [2/4] Checking PyInstaller...
pip show pyinstaller >nul 2>&1
if errorlevel 1 (
    echo PyInstaller not found. Installing PyInstaller...
    pip install pyinstaller
    if errorlevel 1 (
        echo [ERROR] Failed to install PyInstaller
        pause
        exit /b 1
    )
    echo PyInstaller installed successfully
) else (
    echo PyInstaller is already installed
)
echo.

REM Install/verify all dependencies
echo [3/4] Verifying dependencies...
pip install -r requirements.txt
if errorlevel 1 (
    echo [WARNING] Some dependencies may have failed to install
    echo Continuing with build anyway...
)
echo.

REM Clean previous build
if exist "build" (
    echo Cleaning previous build folder...
    rmdir /s /q build
)
if exist "dist\R2_Sequence_Extractor.exe" (
    echo Removing previous executable...
    del /q "dist\R2_Sequence_Extractor.exe"
)
echo.

REM Build the executable
echo [4/4] Building executable with PyInstaller...
echo This may take a few minutes...
echo.
pyinstaller main.spec
if errorlevel 1 (
    echo.
    echo [ERROR] Build failed!
    echo Check the error messages above for details.
    pause
    exit /b 1
)

echo.
echo ============================================================================
echo BUILD COMPLETED SUCCESSFULLY!
echo ============================================================================
echo.
echo Executable location: dist\R2_Sequence_Extractor.exe
echo.
echo You can now:
echo   1. Copy the .exe to any Windows PC (no Python installation needed)
echo   2. Place L5X files in the same folder as the .exe
echo   3. Run R2_Sequence_Extractor.exe to process files
echo.
echo Note: The 'build' folder can be safely deleted if needed.
echo ============================================================================
echo.

REM Open dist folder
if exist "dist\R2_Sequence_Extractor.exe" (
    explorer dist
)

pause
