#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Build script for creating a standalone executable for Moxy Rates Template Transfer

This script uses PyInstaller to create a standalone Windows executable.
"""

import os
import sys
import shutil
import subprocess

def build_executable():
    """Build the executable using PyInstaller."""
    print("Building Moxy Rates Template Transfer executable...")
    
    # Make sure PyInstaller is installed
    try:
        import PyInstaller
    except ImportError:
        print("PyInstaller not found. Installing...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
    
    # Check if icon file exists, if not, create a default one
    icon_file = "app_icon.ico"
    if not os.path.exists(icon_file):
        print(f"Icon file {icon_file} not found. Using default system icon.")
        icon_param = ""
    else:
        icon_param = f"--icon={icon_file}"
    
    # Build command
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--onefile",
        "--windowed",
        icon_param,
        "--version-file=version_info.txt",
        "--name=Moxy Rates Template Transfer",
        "main.py"
    ]
    
    # Remove empty string elements from the command
    cmd = [item for item in cmd if item]
    
    print(f"Running command: {' '.join(cmd)}")
    
    # Run PyInstaller
    subprocess.check_call(cmd)
    
    # Copy necessary files to dist directory
    print("Copying additional files...")
    dist_dir = os.path.join(os.getcwd(), "dist")
    
    files_to_copy = [
        "README.md",
        "requirements.txt"
    ]
    
    for file in files_to_copy:
        if os.path.exists(file):
            shutil.copy2(file, dist_dir)
            print(f"Copied {file} to {dist_dir}")
    
    # Create empty logs directory in dist
    logs_dir = os.path.join(dist_dir, "logs")
    os.makedirs(logs_dir, exist_ok=True)
    print(f"Created logs directory: {logs_dir}")
    
    print("\nBuild completed successfully!")
    print(f"Executable created in: {dist_dir}")
    print("File: Moxy Rates Template Transfer.exe")

if __name__ == "__main__":
    try:
        build_executable()
    except Exception as e:
        print(f"Error building executable: {str(e)}")
        sys.exit(1) 