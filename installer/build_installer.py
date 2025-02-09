"""
Script to build standalone installer for Monte Carlo Excel Add-in.
"""
import os
import sys
import subprocess
from pathlib import Path

def build_executable():
    """Build the standalone executable installer."""
    try:
        # Get the path to PyInstaller
        pyinstaller_path = Path(sys.executable).parent / "Scripts" / "pyinstaller.exe"
        if not pyinstaller_path.exists():
            pyinstaller_path = Path(os.path.expanduser("~")) / "AppData" / "Roaming" / "Python" / "Python312" / "Scripts" / "pyinstaller.exe"
        
        if not pyinstaller_path.exists():
            raise FileNotFoundError("Could not find pyinstaller.exe")
        
        # Create spec file for PyInstaller
        spec_content = """
# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['install.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('../requirements.txt', '.'),
        ('../src/excel/*.xlam', 'excel'),
        ('../LICENSE', '.'),
    ],
    hiddenimports=['xlwings'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='MonteCarloExcel_Setup',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='installer_icon.ico',
    uac_admin=True,
)
"""
        
        # Write spec file
        with open('MonteCarloExcel.spec', 'w') as f:
            f.write(spec_content)
        
        # Build the executable
        subprocess.check_call([
            str(pyinstaller_path),
            "--clean",
            "MonteCarloExcel.spec"
        ])
        
        print("Installer built successfully!")
        print(f"Executable located at: {Path.cwd() / 'dist' / 'MonteCarloExcel_Setup.exe'}")
        
    except Exception as e:
        print(f"Error building installer: {e}")
        sys.exit(1)

if __name__ == "__main__":
    build_executable()
