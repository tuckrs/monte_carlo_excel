import subprocess
import sys
import os
import logging
from pathlib import Path

# Setup logging
logging.basicConfig(
    filename='monte_carlo_install.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

def install_package(package):
    """Install a Python package using pip."""
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", package])
        logger.info(f"Successfully installed {package}")
        return True
    except subprocess.CalledProcessError as e:
        logger.error(f"Failed to install {package}: {e}")
        return False

def main():
    """Main installation function."""
    logger.info("Starting Python dependencies installation")
    
    # Required packages
    packages = [
        "xlwings>=0.30.0",
        "numpy>=1.24.0",
        "pandas>=2.0.0",
        "scipy>=1.10.0"
    ]
    
    # Install each package
    for package in packages:
        if not install_package(package):
            logger.error(f"Installation failed for {package}")
            sys.exit(1)
    
    logger.info("All Python dependencies installed successfully")
    return 0

if __name__ == "__main__":
    sys.exit(main())
