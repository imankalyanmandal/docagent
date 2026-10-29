@echo off
REM Move to the folder this .bat file lives in, so it works no matter
REM where it's double-clicked from.
cd /d "%~dp0"

echo Installing dependencies...
pip install -r requirements.txt
if errorlevel 1 (
    echo.
    echo pip install failed - see the error above.
    pause
    exit /b 1
)

echo.
echo Starting the app...
echo Open http://localhost:5000 in your browser.
echo Press Ctrl+C in this window to stop the server.
echo.
python app.py

pause
