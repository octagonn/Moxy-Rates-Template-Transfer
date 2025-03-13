@echo off
echo Running Moxy Rates Template Transfer Tests...
python test_app.py
if errorlevel 1 (
    echo Error running tests.
    echo Please check that Python and all required dependencies are installed.
    echo Install dependencies with: pip install -r requirements.txt
    pause
)
pause 