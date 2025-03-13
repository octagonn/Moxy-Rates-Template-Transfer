@echo off
echo Moxy Rates Template Transfer - Setup
echo ====================================
echo.

REM Attempt to find Python installation
SET PYTHON_CMD=
FOR %%p IN (python py) DO (
    %%p --version >nul 2>&1
    IF NOT ERRORLEVEL 1 (
        SET PYTHON_CMD=%%p
        GOTO :PYTHON_FOUND
    )
)

:PYTHON_NOT_FOUND
echo Python not found in PATH.
echo Please install Python and make sure it's in your PATH.
echo Visit https://www.python.org/downloads/ to download Python.
echo.
echo After installing Python, run this setup script again.
pause
exit /b 1

:PYTHON_FOUND
%PYTHON_CMD% --version
echo Python found: %PYTHON_CMD%
echo.

echo Installing required dependencies...
%PYTHON_CMD% -m pip install -r requirements.txt
IF ERRORLEVEL 1 (
    echo.
    echo Failed to install dependencies.
    echo Please make sure you have an internet connection and pip is working.
    pause
    exit /b 1
)

echo.
echo Dependencies installed successfully!
echo.
echo Setup complete! You can now run the application using run_app.bat
echo.
pause 