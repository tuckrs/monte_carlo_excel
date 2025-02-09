# Monte Carlo Excel Add-in Installer

This directory contains the installation scripts for the Monte Carlo Excel Add-in.

## Components

1. `install.py` - Python installation script that:
   - Installs required Python packages
   - Registers the Excel add-in
   - Creates necessary startup scripts

2. `monte_carlo_installer.nsi` - NSIS installer script that creates a Windows installer

## Building the Installer

### Prerequisites
1. Install NSIS (Nullsoft Scriptable Install System)
2. Python 3.8+ with pip

### Steps to Build

1. Build Python package:
```bash
cd ..
python setup.py sdist bdist_wheel
```

2. Build Windows installer:
```bash
makensis monte_carlo_installer.nsi
```

This will create `MonteCarloExcelSetup.exe` in the current directory.

## Installation

### Automated Installation
1. Run `MonteCarloExcelSetup.exe`
2. Follow the installation wizard
3. Restart Excel

### Manual Installation
1. Run `install.py` with administrator privileges:
```bash
python install.py
```

2. Restart Excel

## Post-Installation
- The add-in will appear in Excel's ribbon
- All Monte Carlo simulation features will be available
- Check Excel's Add-ins menu to verify installation

## Troubleshooting
1. If the add-in doesn't appear in Excel:
   - Check Excel's "Disabled Items" list
   - Verify registry entries
   - Run Excel as administrator once

2. If Python packages fail to install:
   - Check internet connection
   - Verify Python path
   - Run pip with --verbose flag

## Uninstallation
1. Use Windows Control Panel > Programs and Features, or
2. Run the uninstaller from Start Menu
