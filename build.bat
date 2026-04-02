@echo off
echo Installing dependencies...
pip install -r requirements.txt pyinstaller

echo.
echo Building executable...
pyinstaller --onefile --windowed --name "ClaudeUsageMonitor" --icon=NONE main.py

echo.
echo Done! Executable is in dist\ClaudeUsageMonitor.exe
pause
