@echo off
echo Updating tournament participants...
cd /d "%~dp0\.."
python3 "code\azuriraj_ucesnike.py" "%1"
echo.
echo Update complete! Press any key to close...
pause > nul
