@echo off
echo Building StudioVision Monitor...
echo.

:: Ensure PyInstaller is available
pip install pyinstaller

echo.
echo Compiling — this may take a minute...
pyinstaller --noconsole --onefile --clean --icon=Studiov2000.ico studiovision_monitor_AL.py

echo.
echo Build complete. Output: dist\studiovision_monitor_AL.exe
pause