#!/usr/bin/env python3
"""
Setup script for Email Generator GUI Application
Creates executable and handles dependencies
"""

import os
import sys
import subprocess
import platform

def install_requirements():
    """Install required packages"""
    requirements = [
        'tkinter',  # Usually built-in with Python
        'pandas',
        'requests',
        'pyinstaller'  # For creating executable
    ]
    
    print("Installing required packages...")
    for package in requirements:
        if package == 'tkinter':
            try:
                import tkinter
                print(f"✓ {package} already available")
            except ImportError:
                print(f"✗ {package} not available - install Python with tkinter support")
        else:
            try:
                subprocess.check_call([sys.executable, '-m', 'pip', 'install', package])
                print(f"✓ {package} installed")
            except subprocess.CalledProcessError:
                print(f"✗ Failed to install {package}")

def create_spec_file():
    """Create PyInstaller spec file for better control"""
    spec_content = '''
# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['main_gui.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=['pandas', 'tkinter', 'tkinter.ttk', 'tkinter.filedialog', 'tkinter.messagebox', 'tkinter.scrolledtext'],
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
    name='FreshStartEmailGenerator',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # No console window
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icon.ico' if os.path.exists('icon.ico') else None,
)

# For macOS, create .app bundle
if platform.system() == 'Darwin':
    app = BUNDLE(
        exe,
        name='Fresh Start Email Generator.app',
        icon='icon.icns' if os.path.exists('icon.icns') else None,
        bundle_identifier='com.freshstart.emailgenerator',
    )
'''
    
    with open('email_generator.spec', 'w') as f:
        f.write(spec_content)
    print("✓ Created PyInstaller spec file")

def build_executable():
    """Build the executable"""
    print("Building executable...")
    
    # Create spec file
    create_spec_file()
    
    try:
        # Build using spec file
        subprocess.check_call([
            sys.executable, '-m', 'PyInstaller',
            '--clean',
            'email_generator.spec'
        ])
        print("✓ Executable built successfully!")
        
        # Show output location
        system = platform.system()
        if system == 'Darwin':  # macOS
            print("📁 Executable location: dist/Fresh Start Email Generator.app")
        elif system == 'Windows':
            print("📁 Executable location: dist/FreshStartEmailGenerator.exe")
        else:  # Linux
            print("📁 Executable location: dist/FreshStartEmailGenerator")
            
    except subprocess.CalledProcessError as e:
        print(f"✗ Failed to build executable: {e}")

def create_requirements_file():
    """Create requirements.txt file"""
    requirements = """pandas>=1.3.0
requests>=2.25.0
pyinstaller>=4.0
"""
    
    with open('requirements.txt', 'w') as f:
        f.write(requirements)
    print("✓ Created requirements.txt")

def create_readme():
    """Create README file"""
    readme_content = """# Fresh Start Cleaning Email Generator

A GUI application for generating and sending personalized emails to potential cleaning service clients.

## Features

- 📊 CSV upload for prospect data
- 🤖 AI-powered email generation using Ollama
- ✏️ Email editing before sending
- 📧 Automated email sending via Gmail
- 📁 Configuration management
- 📈 Email history and reporting

## Setup Instructions

### 1. Install Ollama
Download and install Ollama from: https://ollama.ai

### 2. Install AI Model
```bash
ollama pull mistral
```

### 3. Gmail App Password
1. Go to Google Account → Security → 2-Step Verification
2. App passwords → Generate new password
3. Copy the 16-character password

### 4. Run Application

#### Option A: Run Python Script
```bash
python main_gui.py
```

#### Option B: Use Executable
Double-click the executable file:
- **macOS**: Fresh Start Email Generator.app
- **Windows**: FreshStartEmailGenerator.exe
- **Linux**: FreshStartEmailGenerator

## CSV Format

Your CSV file should include these columns:
- Company Name (required)
- Email (required)
- Industry
- Contact Name
- Company Size
- Location
- Notes

## Usage

1. **Configure Settings**: Set up email and company info in Configuration tab
2. **Upload CSV**: Load your prospects data
3. **Generate Emails**: AI creates personalized emails
4. **Review & Edit**: Modify emails before sending
5. **Send**: Send individual emails or batch send

## Files Created

- `config.json`: Stores your settings
- `email_results.json`: Email history
- `prospects_template.csv`: Sample CSV format

## Support

For issues or questions, contact Fresh Start Cleaning Co.
Website: https://freshcleaningcolouisiana.com/
"""

    with open('README.md', 'w') as f:
        f.write(readme_content)
    print("✓ Created README.md")

def create_batch_files():
    """Create convenience batch/shell files"""
    
    # Windows batch file
    if platform.system() == 'Windows':
        batch_content = """@echo off
echo Starting Fresh Start Email Generator...
python main_gui.py
pause
"""
        with open('run_email_generator.bat', 'w') as f:
            f.write(batch_content)
        print("✓ Created run_email_generator.bat")
    
    # macOS/Linux shell script
    else:
        shell_content = """#!/bin/bash
echo "Starting Fresh Start Email Generator..."
python3 main_gui.py
"""
        with open('run_email_generator.sh', 'w') as f:
            f.write(shell_content)
        os.chmod('run_email_generator.sh', 0o755)  # Make executable
        print("✓ Created run_email_generator.sh")

def main():
    """Main setup function"""
    print("🚀 Fresh Start Cleaning Email Generator Setup")
    print("=" * 50)
    
    # Check Python version
    if sys.version_info < (3, 7):
        print("❌ Python 3.7 or higher required")
        sys.exit(1)
    
    print(f"✓ Python {sys.version_info.major}.{sys.version_info.minor} detected")
    
    # Create files
    create_requirements_file()
    create_readme()
    create_batch_files()
    
    # Install requirements
    install_requirements()
    
    # Ask about building executable
    response = input("\n🔨 Build executable? (y/n): ").strip().lower()
    if response == 'y':
        build_executable()
    
    print("\n✅ Setup complete!")
    print("\nNext steps:")
    print("1. Install Ollama: https://ollama.ai")
    print("2. Run: ollama pull mistral")
    print("3. Run the application:")
    
    if os.path.exists('dist'):
        system = platform.system()
        if system == 'Darwin':
            print("   - Double-click: dist/Fresh Start Email Generator.app")
        elif system == 'Windows':
            print("   - Double-click: dist/FreshStartEmailGenerator.exe")
        else:
            print("   - Run: ./dist/FreshStartEmailGenerator")
    
    print("   - Or run: python main_gui.py")

if __name__ == "__main__":
    main()