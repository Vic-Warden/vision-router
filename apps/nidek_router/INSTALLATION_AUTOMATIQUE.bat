@echo off
title Box 2 Image Router - Installation
cls

:: Change working directory to the script's own folder
cd /d "%~dp0"

:: 1. AUTO-ELEVATION — Re-launch with admin rights if not already elevated
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo Requesting administrator privileges...
    powershell -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
    exit /b
)


echo   BOX 2 IMAGE ROUTER - INSTALLATION

echo.

:: 2. SILENT INSTALL — Microsoft Access Database Engine (x64)
echo [1/3] Installing Microsoft Access Database Engine...
if exist "accessdatabaseengine_X64.exe" (
    start /wait "" "accessdatabaseengine_X64.exe" /quiet
    echo [OK] Access Database Engine installed.
) else (
    echo [SKIP] accessdatabaseengine_X64.exe not found - step skipped.
)
echo.

:: 3. DEPLOYMENT — Copy application files to target directory
echo [2/3] Deploying application files...
set "TARGET_DIR=C:\Routeur_Images_Box2"
if not exist "%TARGET_DIR%" mkdir "%TARGET_DIR%"

if exist "box2V4_AL.exe" (
    copy /y "box2V4_AL.exe" "%TARGET_DIR%\box2V4_AL.exe" >nul
    echo [OK] Executable copied to: %TARGET_DIR%
) else (
    echo [ERROR] box2V4_AL.exe not found.
    echo Make sure all files are in the same folder as this installer.
    pause
    exit /b
)

if exist "Studiov2000.ico" (
    copy /y "Studiov2000.ico" "%TARGET_DIR%\Studiov2000.ico" >nul
    echo [OK] Icon copied to: %TARGET_DIR%
) else (
    echo [SKIP] Studiov2000.ico not found - icon step skipped.
)
echo.

:: 4. FIRST-RUN SETUP — Launch the application to trigger configuration wizard
echo [3/3] Launching first-run configuration...
echo.

echo   FILES INSTALLED SUCCESSFULLY

echo.
echo Follow the on-screen instructions to complete the setup.
echo.

start "" "%TARGET_DIR%\box2V4_AL.exe"
exit