@echo off

set PYTHON_VERSION=3.7.7

REM Check if Python is already installed
python --version > nul 2>&1
if %errorlevel% equ 0 (
  echo Python is already installed.
  echo Sleeping for 5 seconds...
@REM   timeout /t 5
@REM   goto :eof
) else (
    if exist "%ProgramFiles(x86)%" (
      set PYTHON_URL=https://www.python.org/ftp/python/%PYTHON_VERSION%/python-%PYTHON_VERSION%.exe
    ) else (
      set PYTHON_URL=https://www.python.org/ftp/python/%PYTHON_VERSION%/python-%PYTHON_VERSION%-amd64.exe
    )

    echo Downloading Python %PYTHON_VERSION% from %PYTHON_URL%...

    curl -o python.exe %PYTHON_URL%

    echo Installing Python %PYTHON_VERSION%...

    start /wait python.exe /quiet InstallAllUsers=1 PrependPath=1 Include_test=0

    echo Setting Python environment variable...

    set "PYTHON_INSTALL_DIR=C:\Python%PYTHON_VERSION%"
    set "PYTHON_SCRIPTS_DIR=C:\Python%PYTHON_VERSION%\Scripts"
    setx PATH "%PATH%;%PYTHON_INSTALL_DIR%;%PYTHON_SCRIPTS_DIR%" /M

    # setx PATH "%PATH%;C:\Python%PYTHON_VERSION%;C:\Python%PYTHON_VERSION%\Scripts" /M

    echo Cleaning up...

    del python.exe

    echo Python installation complete.
)
cd "%~dp0"\orderExtractor\

echo installing dependencies
pip install -r requirements.txt

echo Running GUI
python gui.py
