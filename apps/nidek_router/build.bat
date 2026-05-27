@echo off
echo Building Box 2 Image Router...
echo.

:: Ensure PyInstaller is available
pip install pyinstaller

echo.
echo Compiling — this may take a minute...
pyinstaller --noconsole --onefile --clean --icon=Studiov2000.ico box2V4_AL.py

echo.
echo Build complete. Output: dist\box2V4_AL.exe
pause