@echo off
echo Building Moxy Rates Template Transfer executable...
python build_exe.py
if errorlevel 1 (
    echo Error building executable.
    echo Please check that Python and all required dependencies are installed.
    echo Install dependencies with: pip install -r requirements.txt
    pause
) else (
    echo Build completed successfully!
    echo Executable is in the 'dist' folder.
)
pause 