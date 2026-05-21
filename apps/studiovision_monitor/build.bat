@echo off
echo   Compiling StudioVision Monitor
echo.

echo Checking for PyInstaller...
pip install pyinstaller

echo.
echo Compiling in progress (this may take a minute)...
pyinstaller --noconsole --onefile --clean studiovision_monitor_AL.py

echo.
echo Done! Your .exe application is in the "dist" folder.
pause