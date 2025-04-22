import os
import subprocess
import shutil

def build_executable():
    print("Starting build process for BSCScan Automation Tool...")
    
    # Create spec file content
    spec_content = """# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(['bscscan_app.py'],
             pathex=[],
             binaries=[],
             datas=[],
             hiddenimports=['pkg_resources.py2_warn'],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher,
             noarchive=False)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,
          [],
          name='BSCScan Automation',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          upx_exclude=[],
          runtime_tmpdir=None,
          console=False)
"""
    
    # Write spec file
    with open("bscscan_app.spec", "w") as f:
        f.write(spec_content)
    
    print("Created spec file.")
    
    # Run PyInstaller
    print("Running PyInstaller (this may take several minutes)...")
    subprocess.run(["pyinstaller", "--clean", "bscscan_app.spec"], check=True)
    
    print("Build complete!")
    print("Executable created in the 'dist' folder.")
    
    # Create distribution package
    dist_dir = "BSCScan Automation Package"
    if os.path.exists(dist_dir):
        shutil.rmtree(dist_dir)
    os.makedirs(dist_dir)
    
    # Copy executable
    shutil.copy("dist/BSCScan Automation.exe", dist_dir)
    
    # Create README file
    readme_content = """BSCScan Automation Tool
======================

This tool automates form filling on BSCScan's contract interface using data from Google Sheets.

Setup Instructions:
------------------
1. Create a Google Service Account and download the JSON key file
   - Go to https://console.cloud.google.com/
   - Create a new project
   - Enable the Google Sheets API
   - Create a service account
   - Download the JSON key file

2. Share your Google Sheet with the service account email address

Usage:
-----
1. Launch the application
2. Enter your Google Sheet URL
3. Set the start and end rows
4. Browse for your service account JSON file
5. Click "Preview Data" to check your data
6. Click "Start Automation" to begin
7. Complete any CAPTCHA or verification in the browser when prompted
8. Monitor progress in the status window

Support:
-------
For help or issues, contact the developer.
"""
    
    with open(os.path.join(dist_dir, "README.txt"), "w") as f:
        f.write(readme_content)
    
    print(f"Distribution package created in '{dist_dir}' folder.")

if __name__ == "__main__":
    build_executable()