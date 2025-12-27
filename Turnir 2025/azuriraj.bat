@echo off
echo.
echo Updating participants for tournament: %~dp0
echo.

REM Save current directory and change to the parent directory (project root)
pushd "%~dp0\.."

REM Run the Python script with current folder as parameter
REM Remove trailing backslash from %~dp0 to avoid quote issues
set "TOURNAMENT_FOLDER=%~dp0"
set "TOURNAMENT_FOLDER=%TOURNAMENT_FOLDER:~0,-1%"
python3 "code\azuriraj_ucesnike.py" "%TOURNAMENT_FOLDER%"

REM Return to original directory
popd


