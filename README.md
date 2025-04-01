# Moxy Rates Template Transfer

A Windows application that transfers data from Adjusted Rates Excel files to Template Excel files with smart column mapping.

## Key Features

- Easy-to-use interface for selecting files
- Automatic file format detection
- Smart column mapping for different naming conventions
- Template integration for consistent output
- Save and reuse mapping templates

## Installation

### Quick Install (Recommended)
1. Download the installer from [Releases](https://github.com/your-username/moxy-rates-template-transfer/releases)
2. Run the installer and follow the instructions
3. Launch from Start menu or desktop shortcut

### Run from Source (Windows)
1. Install Python 3.9+ with "Add Python to PATH" enabled
2. Clone or download this repository
3. Run `setup.bat` to install dependencies
4. Start the app with `run_app.bat`

## How to Use

1. Launch the application
2. Select your Adjusted Rates Excel file
3. Select your Template Excel file
4. Set an output filename or use the default
5. Click "Process Files"

## Advanced Options

- **Preview Mapping**: View and adjust column mapping before processing
- **Auto-detect**: Automatically identify file structures
- **Use saved mappings**: Apply previously saved mappings
- **Save mapping templates**: Save current mapping for future use

## Required Columns

The Adjusted Rates file should include these columns (names may vary):
- Coverage
- Term
- Miles (or FromMiles/ToMiles)
- MinYear
- MaxYears
- Class
- RateCost
- Deductible

## Troubleshooting

- If column mapping fails, try manual mapping
- Ensure your source file has all required data
- Close Excel files before processing
- Check logs in the `logs` folder for details

## License

This project is licensed under the MIT License. 