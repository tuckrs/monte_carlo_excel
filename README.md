# Monte Carlo Excel Add-in

A powerful Monte Carlo simulation engine implemented as an Excel add-in, providing advanced statistical analysis and visualization capabilities.

## Features

- Monte Carlo simulation with normal distribution sampling
- Statistical analysis including mean, standard deviation, and percentiles
- Custom Excel ribbon interface for easy access
- High-performance Python backend using numpy and pandas
- Seamless Excel integration through xlwings

## Installation

### Prerequisites
- Windows 10 or later
- Microsoft Excel 2016 or later
- Python 3.12 (will be prompted to install if not present)

### Installation Steps

1. Download the latest release:
   - Go to the [Releases](https://github.com/tuckrs/monte_carlo_excel/releases) page
   - Download `MonteCarloExcelSetup.zip`
   - Extract the zip file

2. Run the installer:
   - Double-click `MonteCarloExcelSetup.exe`
   - If Python 3.12 is not installed, the installer will direct you to install it
   - Follow the installation wizard

3. After installation:
   - Open Excel
   - You should see a new "Monte Carlo" tab in the ribbon
   - If you don't see the tab, go to File > Options > Add-ins > Manage Excel Add-ins > Go... and ensure "Monte Carlo Excel" is checked

## Usage

1. Enter your data:
   - Organize your data in columns with headers
   - Each column will be treated as a separate variable

2. Run a simulation:
   - Select your data range (including headers)
   - Click the "Run Simulation" button in the Monte Carlo tab
   - Select where to put the results
   - The add-in will generate statistics including:
     - Mean
     - Standard Deviation
     - 5th and 95th percentiles

3. Need help?
   - Click the "Help" button in the Monte Carlo tab for quick instructions

## Project Structure

```
monte_carlo_excel/
├── src/                    # Source code
│   └── excel/             # Excel add-in files
├── installer/             # Installer scripts
│   ├── monte_carlo_installer.nsi    # NSIS installer script
│   ├── install_dependencies.py      # Python package installer
│   └── install_excel_addin.ps1     # PowerShell add-in installer
└── docs/                  # Documentation
```

## Development

### Requirements
- NSIS (Nullsoft Scriptable Install System)
- Python 3.12
- Microsoft Excel
- Required Python packages:
  - xlwings
  - numpy
  - pandas
  - scipy

### Building the Installer

1. Install NSIS from https://nsis.sourceforge.io/Download
2. Run:
   ```bash
   cd installer
   makensis monte_carlo_installer.nsi
   ```
3. The installer will be created as `MonteCarloExcelSetup.exe`

## License

This software is proprietary and confidential. Unauthorized copying, distribution, or use is strictly prohibited.

## Support

For issues, feature requests, or questions, please create an issue in the GitHub repository.
