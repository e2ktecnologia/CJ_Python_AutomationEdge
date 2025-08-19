@echo off

:: Check for Python installation
where python >nul 2>nul
if errorlevel 1 (
    echo Python is not installed. Please install Python and try again.
    exit /b
)

:: Create a virtual environment in the current directory if not already done
if not exist "%cd%\venv" (
    python -m venv "%cd%\venv"
    echo Virtual environment created in the current directory.
) else (
    echo Virtual environment already exists in the current directory.
)

:: Activate the virtual environment
call "%cd%\venv\Scripts\activate"

:: Upgrade pip
python -m pip install --upgrade pip

:: Install required libraries
pip install clicknium pandas pywin32 requests

:: Inform user of completion
echo All libraries have been installed.

deactivate

:: Pause to allow the user to review the results
pause
