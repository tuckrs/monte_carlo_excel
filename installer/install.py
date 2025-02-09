"""
Installation script for Monte Carlo Excel Add-in.
"""
import os
import sys
import site
import winreg
import subprocess
import logging
from pathlib import Path
import shutil
import xlwings

# Setup logging
logging.basicConfig(
    filename='monte_carlo_install.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

def is_admin():
    """Check if script is running with admin privileges."""
    try:
        return os.getuid() == 0
    except AttributeError:
        import ctypes
        return ctypes.windll.shell32.IsUserAnAdmin() != 0

def check_prerequisites():
    """Check if all required files and dependencies are present."""
    logger.info("Checking prerequisites...")
    required_files = ['requirements.txt', 'monte_carlo.xlam']
    
    for file in required_files:
        if not os.path.exists(file):
            logger.error(f"Required file {file} not found!")
            return False
    return True

def install_dependencies():
    """Install required Python packages."""
    logger.info("Installing Python dependencies...")
    try:
        # Get Python executable path
        python_path = sys.executable
        logger.info(f"Using Python from: {python_path}")
        
        requirements_path = os.path.join(os.getcwd(), "requirements.txt")
        logger.info(f"Requirements file path: {requirements_path}")
        
        if not os.path.exists(requirements_path):
            logger.error("requirements.txt not found!")
            return False
            
        result = subprocess.run(
            [python_path, "-m", "pip", "install", "-r", requirements_path],
            capture_output=True,
            text=True
        )
        
        if result.returncode != 0:
            logger.error(f"Failed to install dependencies: {result.stderr}")
            return False
            
        logger.info("Dependencies installed successfully")
        return True
    except Exception as e:
        logger.error(f"Error installing dependencies: {str(e)}")
        return False

def register_excel_addin():
    """Register the Excel add-in."""
    logger.info("Registering Excel add-in...")
    try:
        # Get the xlwings addin path
        xlwings_addin = Path(xlwings.__path__[0]) / "addin" / "xlwings.xlam"
        
        # Copy our custom ribbon UI file
        addin_dir = Path(site.getsitepackages()[0]) / "monte_carlo_excel" / "excel"
        ribbon_path = addin_dir / "monte_carlo.xlam"
        
        # Register add-in with Excel
        excel_key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Office\Excel\Addins\MonteCarloExcel",
            0,
            winreg.KEY_WRITE
        )
        
        winreg.SetValueEx(
            excel_key,
            "Description",
            0,
            winreg.REG_SZ,
            "Monte Carlo Simulation Add-in for Excel"
        )
        winreg.SetValueEx(
            excel_key,
            "FriendlyName",
            0,
            winreg.REG_SZ,
            "Monte Carlo Excel"
        )
        winreg.SetValueEx(
            excel_key,
            "LoadBehavior",
            0,
            winreg.REG_DWORD,
            3
        )
        winreg.SetValueEx(
            excel_key,
            "Path",
            0,
            winreg.REG_SZ,
            str(ribbon_path)
        )
        
        return True
    except Exception as e:
        logger.error(f"Failed to register Excel add-in: {e}")
        return False

def create_startup_script():
    """Create Excel startup script for the add-in."""
    logger.info("Creating startup script...")
    try:
        startup_dir = Path(os.getenv('APPDATA')) / "Microsoft" / "Excel" / "XLSTART"
        startup_dir.mkdir(parents=True, exist_ok=True)
        
        script_content = """
from monte_carlo_excel.excel.addin import MonteCarloAddin

def monte_carlo_init():
    addin = MonteCarloAddin()
    addin.setup_ribbon()
"""
        
        startup_script = startup_dir / "monte_carlo_startup.py"
        with open(startup_script, 'w') as f:
            f.write(script_content)
            
        return True
    except Exception as e:
        logger.error(f"Failed to create startup script: {e}")
        return False

def main():
    """Main installation function."""
    logger.info("Starting Monte Carlo Excel Add-in installation...")
    
    if not is_admin():
        logger.error("Installation requires admin privileges!")
        sys.exit(1)
        
    if not check_prerequisites():
        logger.error("Prerequisites check failed!")
        sys.exit(1)
        
    if not install_dependencies():
        logger.error("Failed to install dependencies!")
        sys.exit(1)
        
    try:
        register_excel_addin()
        create_startup_script()
        logger.info("Installation completed successfully!")
    except Exception as e:
        logger.error(f"Installation failed: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    main()
