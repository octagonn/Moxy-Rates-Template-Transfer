# Moxy Rates Template Transfer

A Windows application that automates the process of transferring data from Adjusted Rates Excel spreadsheets into Template Excel spreadsheets, with intelligent column mapping and format detection.

## Features

- **Intuitive GUI Interface**: Easy-to-use interface for file selection and configuration
- **Intelligent File Format Detection**: Automatically analyzes and identifies file structures
- **Smart Column Mapping**: Handles different column naming conventions across files
- **Template Integration**: Seamlessly integrates data with existing templates
- **Mapping Templates**: Save and reuse successful mappings for similar files
- **Visual Mapping Interface**: Manual mapping interface when automatic detection is uncertain

## Installation

### Option 1: Using the Installer (Recommended)

1. Download the latest installer from the [Releases](https://github.com/your-username/moxy-rates-template-transfer/releases) page
2. Run the installer and follow the on-screen instructions
3. Launch the application from the Start menu or desktop shortcut

### Option 2: Running from Source (Windows)

1. Make sure you have Python 3.9+ installed with "Add Python to PATH" enabled during installation
2. Clone or download this repository
3. **Simple setup:** 
   - Double-click `setup.bat` to install all dependencies automatically
   - After setup completes, run `run_app.bat` to start the application

4. **Manual setup:** If the simple setup doesn't work
   - Open Command Prompt as administrator
   - Navigate to the repository folder: `cd path\to\moxy-rates-template-transfer`
   - Install the required dependencies: `python -m pip install -r requirements.txt`
   - Run the application: `pythonw main.pyw` or use `run_app.bat`

### Troubleshooting Setup

- **Python not found**: Make sure Python is installed and added to your PATH
- **pip not recognized**: Try using `python -m pip` instead of just `pip`
- **Import errors**: Make sure all dependencies are installed with `python -m pip install -r requirements.txt`

## Usage

### Basic Usage

1. Launch the application
2. Click **Browse...** to select your Adjusted Rates Excel file
3. Click **Browse...** to select your Template Excel file
4. Specify an output filename or use the default
5. Click **Process Files** to start the transfer

### Advanced Features

- **Preview Mapping**: Click this button to see and modify the column mapping before processing
- **Auto-detect file formats**: Enable to automatically detect file structures
- **Use saved mappings**: Enable to reuse previously saved mappings for similar files
- **Save mapping templates**: In the mapping dialog, check "Save as mapping template" to reuse mappings

## File Requirements

### Adjusted Rates File

The Adjusted Rates file should contain the following columns (though the exact names can vary):

- Coverage
- Term
- Miles (or FromMiles/ToMiles)
- MinYear
- MaxYears
- Class
- RateCost
- Deductible

### Template File

Any Excel file can be used as a template. The application will:

1. Detect the structure of the template
2. Add any missing columns required for the data
3. Transfer the transformed data into the template format

## Troubleshooting

### Common Issues

- **File Format Not Detected**: Try disabling "Auto-detect file formats" and manually map columns
- **Missing Required Columns**: Ensure your Adjusted Rates file contains all required data fields
- **Excel File Access Error**: Close the file in Excel before processing

### Logs

Application logs are stored in the `logs` folder in the application directory. These can be helpful for troubleshooting issues.

## Building from Source

To create a standalone executable:

1. Install PyInstaller:
   ```
   pip install pyinstaller
   ```

2. Create the executable:
   ```
   pyinstaller --onefile --windowed --icon=app_icon.ico --name="Moxy Rates Template Transfer" main.py
   ```

3. The executable will be created in the `dist` folder

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the MIT License - see the LICENSE file for details. 