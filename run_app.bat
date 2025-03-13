@echo off
echo Starting Moxy Rates Template Transfer Application...

REM Attempt to find Python installation
SET PYTHON_CMD=
FOR %%p IN (python pythonw py) DO (
    %%p --version >nul 2>&1
    IF NOT ERRORLEVEL 1 (
        SET PYTHON_CMD=%%p
        GOTO :PYTHON_FOUND
    )
)

:PYTHON_NOT_FOUND
echo Python not found in PATH.
echo Please install Python and make sure it's in your PATH, or run this app with your Python installation.
echo Visit https://www.python.org/downloads/ to download Python.
pause
exit /b 1

:PYTHON_FOUND
echo Found Python: %PYTHON_CMD%

REM Check for main.pyw
IF NOT EXIST main.pyw (
    echo Error: main.pyw not found.
    echo Please make sure you're running this batch file from the repository root directory.
    pause
    exit /b 1
)

REM Check for requirements
%PYTHON_CMD% -c "import pandas, openpyxl, configparser, fuzzywuzzy" >nul 2>&1
IF ERRORLEVEL 1 (
    echo Some dependencies are missing. Would you like to install them now? (Y/N)
    SET /P INSTALL_DEPS=
    IF /I "%INSTALL_DEPS%"=="Y" (
        echo Installing dependencies...
        %PYTHON_CMD% -m pip install -r requirements.txt
        IF ERRORLEVEL 1 (
            echo Failed to install dependencies.
            pause
            exit /b 1
        )
    ) ELSE (
        echo Dependencies required to run the application are missing.
        pause
        exit /b 1
    )
)

REM Run the application
echo Starting application...
IF "%PYTHON_CMD%"=="python" (
    start /b pythonw main.pyw
) ELSE (
    start /b %PYTHON_CMD% main.pyw
)

REM Check for immediate errors
timeout /t 1 /nobreak >nul
IF ERRORLEVEL 1 (
    echo Error running application.
    echo Please check the logs for more information.
    pause
) 