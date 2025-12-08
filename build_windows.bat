@echo off
echo Installing dependencies...
pip install -r requirements.txt

echo Building Windows Executable...
pyinstaller --onefile --clean --name filter_bonus_tool filter_bonus_data.py

echo.
echo Build finished!
echo The executable is located in the "dist" folder: dist\filter_bonus_tool.exe
pause
