"""
PyInstaller build script for BoMination App
Enhanced version for sales team distribution
"""
import PyInstaller.__main__
from pathlib import Path
import sys
import subprocess
import os

def check_dependencies():
    """Check if all required dependencies are installed."""
    print("🔍 Checking dependencies...")
    
    # Map package names to their import names
    required_packages = {
        'pyinstaller': 'PyInstaller',
        'pandas': 'pandas',
        'selenium': 'selenium',
        'ttkbootstrap': 'ttkbootstrap',
        'tabula-py': 'tabula',
        'openpyxl': 'openpyxl',
        'matplotlib': 'matplotlib',
        'pandastable': 'pandastable'
    }
    
    missing = []
    for package_name, import_name in required_packages.items():
        try:
            __import__(import_name)
            print(f"✅ {package_name}")
        except ImportError:
            missing.append(package_name)
            print(f"❌ {package_name}")
    
    if missing:
        print(f"\n⚠️  Missing packages: {', '.join(missing)}")
        print("Please install them using: pip install -r requirements.txt")
        return False
    
    print("✅ All dependencies found!")
    return True

def build_app():
    """Build the BoMination app using PyInstaller."""
    
    if not check_dependencies():
        sys.exit(1)
    
    # Get the current directory
    current_dir = Path(__file__).parent
    src_dir = current_dir / "src"
    
    # Main script path
    main_script = str(src_dir / "BoMinationApp.py")
    
    # Chromedriver path
    chromedriver_path = str(src_dir / "chromedriver.exe")
    
    # Files directory path (contains cost sheet template)
    files_dir = current_dir / "Files"
    cost_sheet_path = str(files_dir / "OCTF-1539-COST SHEET.xlsx")
    
    # Verify critical files exist
    critical_files = [main_script, chromedriver_path, cost_sheet_path]
    missing_files = [f for f in critical_files if not Path(f).exists()]
    
    if missing_files:
        print(f"❌ Missing critical files:")
        for f in missing_files:
            print(f"   - {f}")
        sys.exit(1)
    
    print("✅ All critical files found!")
    
    # All Python source files that might be called as subprocesses
    src_files = [
        str(src_dir / "main_pipeline.py"),
        str(src_dir / "extract_bom_tab.py"),
        str(src_dir / "lookup_price.py"),
        str(src_dir / "map_cost_sheet.py"),
        str(src_dir / "validation_utils.py"),
    ]
    
    # Build arguments for PyInstaller
    args = [
        main_script,
        '--name=BoMinationApp',
        '--onefile',  # Create a single executable
        '--windowed',  # Don't show console window
        '--noconfirm',  # Overwrite output directory without confirmation
        f'--add-data={chromedriver_path};.',  # Include chromedriver
        f'--add-data={cost_sheet_path};Files',  # Include cost sheet template
        f'--add-data={src_dir};src',  # Include all source files
        '--hidden-import=ttkbootstrap',
        '--hidden-import=selenium',
        '--hidden-import=pandas',
        '--hidden-import=numpy',
        '--hidden-import=matplotlib',
        '--hidden-import=openpyxl',
        '--hidden-import=xlrd',
        '--hidden-import=tabula',
        '--hidden-import=pandastable',
        '--hidden-import=PIL',
        '--hidden-import=tkinter',
        '--hidden-import=jpype1',
        '--hidden-import=jpype1._jvmfinder',
        '--hidden-import=packaging',
        '--hidden-import=packaging.version',
        '--hidden-import=packaging.specifiers',
        '--hidden-import=packaging.requirements',
        '--collect-submodules=selenium',
        '--collect-submodules=ttkbootstrap',
        '--collect-submodules=pandas',
        '--collect-submodules=tabula',
        '--collect-submodules=packaging',
        '--collect-data=ttkbootstrap',
        '--collect-data=tabula',
        '--additional-hooks-dir=.',  # Include our custom hooks
        '--distpath=dist',
        '--workpath=build',
        '--specpath=.',
        '--exclude-module=matplotlib.tests',
        '--exclude-module=numpy.tests',
        '--exclude-module=pandas.tests',
        '--exclude-module=PIL.tests',
        # Add icon if available
        # '--icon=icon.ico',  # Uncomment if you have an icon file
    ]
    
    print("\n🚀 Starting PyInstaller build...")
    print(f"📄 Main script: {main_script}")
    print(f"🌐 Chromedriver: {chromedriver_path}")
    print(f"📊 Cost sheet: {cost_sheet_path}")
    print(f"📁 Source files: {len(src_files)} files")
    print(f"🎯 Output: dist/BoMinationApp.exe")
    print()
    
    # Run PyInstaller
    try:
        PyInstaller.__main__.run(args)
        print("\n✅ Build completed successfully!")
        print("📦 Check the 'dist' folder for BoMinationApp.exe")
        
        # Check if the executable was created
        exe_path = current_dir / "dist" / "BoMinationApp.exe"
        if exe_path.exists():
            size_mb = exe_path.stat().st_size / (1024 * 1024)
            print(f"📏 Executable size: {size_mb:.1f} MB")
        else:
            print("⚠️  Executable not found - check build logs for errors")
            
    except Exception as e:
        print(f"\n❌ Build failed: {e}")
        sys.exit(1)

if __name__ == "__main__":
    build_app()
