#!/usr/bin/env python3
"""
Diagnostic tool for the sales team to identify ChromeDriver issues.
This script helps diagnose the specific problem when running the BoMination application.
"""

import sys
import os
import subprocess
from pathlib import Path
import tempfile
import platform
import logging

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')
logger = logging.getLogger(__name__)

def check_system_info():
    """Check basic system information."""
    print("="*60)
    print("SYSTEM INFORMATION")
    print("="*60)
    
    print(f"Operating System: {platform.system()} {platform.release()}")
    print(f"Python Version: {sys.version}")
    print(f"Python Executable: {sys.executable}")
    print(f"Current Working Directory: {os.getcwd()}")
    
    # Check if running from PyInstaller
    if getattr(sys, 'frozen', False):
        print("✅ Running from PyInstaller executable")
        print(f"PyInstaller _MEIPASS: {sys._MEIPASS}")
        print(f"Executable location: {sys.executable}")
    else:
        print("✅ Running from Python script")
        print(f"Script location: {__file__}")
    
    print()

def check_chrome_installation():
    """Check if Chrome browser is installed."""
    print("="*60)
    print("CHROME BROWSER CHECK")
    print("="*60)
    
    chrome_paths = [
        "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe",
        "C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe",
        os.path.expanduser("~\\AppData\\Local\\Google\\Chrome\\Application\\chrome.exe")
    ]
    
    chrome_found = False
    for chrome_path in chrome_paths:
        if os.path.exists(chrome_path):
            print(f"✅ Chrome found at: {chrome_path}")
            chrome_found = True
            
            # Get Chrome version
            try:
                result = subprocess.run([chrome_path, "--version"], capture_output=True, text=True, timeout=5)
                if result.returncode == 0:
                    print(f"   Version: {result.stdout.strip()}")
            except Exception as e:
                print(f"   Could not get version: {e}")
            break
    
    if not chrome_found:
        print("❌ Chrome browser not found in common locations")
        print("Please install Google Chrome from: https://www.google.com/chrome/")
    
    print()

def check_chromedriver_locations():
    """Check for ChromeDriver in all possible locations."""
    print("="*60)
    print("CHROMEDRIVER LOCATION CHECK")
    print("="*60)
    
    # Define all possible ChromeDriver locations
    possible_paths = [
        "chromedriver.exe",
        "src/chromedriver.exe",
        str(Path(__file__).parent / "chromedriver.exe"),
        str(Path(__file__).parent / "src" / "chromedriver.exe"),
    ]
    
    # Add PyInstaller-specific paths
    if getattr(sys, 'frozen', False):
        possible_paths.extend([
            str(Path(sys._MEIPASS) / "chromedriver.exe"),
            str(Path(sys._MEIPASS) / "src" / "chromedriver.exe"),
            str(Path(sys.executable).parent / "chromedriver.exe"),
        ])
    
    chromedriver_found = False
    for path in possible_paths:
        exists = os.path.exists(path)
        print(f"  {path} -> {'✅ Found' if exists else '❌ Not found'}")
        
        if exists:
            chromedriver_found = True
            print(f"    File size: {os.path.getsize(path)} bytes")
            
            # Test ChromeDriver
            try:
                result = subprocess.run([path, "--version"], capture_output=True, text=True, timeout=10)
                if result.returncode == 0:
                    print(f"    Version: {result.stdout.strip()}")
                else:
                    print(f"    ❌ Error running ChromeDriver: {result.stderr}")
            except subprocess.TimeoutExpired:
                print(f"    ❌ ChromeDriver timed out")
            except Exception as e:
                print(f"    ❌ Error testing ChromeDriver: {e}")
    
    if not chromedriver_found:
        print("\n❌ ChromeDriver not found in any location!")
        print("Please ensure chromedriver.exe is in one of the checked locations.")
    
    print()

def check_network_connectivity():
    """Check network connectivity to the target website."""
    print("="*60)
    print("NETWORK CONNECTIVITY CHECK")
    print("="*60)
    
    try:
        # Try using urllib instead of requests for better compatibility
        import urllib.request
        import urllib.error
        
        response = urllib.request.urlopen("https://www.oemsecrets.com", timeout=10)
        if response.getcode() == 200:
            print("✅ Successfully connected to oemsecrets.com")
        else:
            print(f"⚠️ Connected to oemsecrets.com but got status code: {response.getcode()}")
    except urllib.error.URLError as e:
        print(f"❌ Could not connect to oemsecrets.com: {e}")
        print("Check your internet connection or firewall settings.")
    except Exception as e:
        print(f"❌ Error checking connectivity: {e}")
    
    print()

def check_antivirus_interference():
    """Check for potential antivirus interference."""
    print("="*60)
    print("ANTIVIRUS INTERFERENCE CHECK")
    print("="*60)
    
    print("Common signs of antivirus interference:")
    print("• ChromeDriver starts but immediately closes")
    print("• Permission denied errors when running ChromeDriver")
    print("• Browser opens but doesn't navigate to websites")
    print("• Slow performance or unexpected timeouts")
    print()
    
    print("If you suspect antivirus interference:")
    print("1. Temporarily disable real-time scanning")
    print("2. Add the application folder to antivirus exclusions")
    print("3. Add chromedriver.exe to the antivirus whitelist")
    print("4. Try running the application as administrator")
    print()

def test_minimal_selenium():
    """Test minimal Selenium functionality."""
    print("="*60)
    print("SELENIUM FUNCTIONALITY TEST")
    print("="*60)
    
    try:
        # Add src directory to path for imports
        if getattr(sys, 'frozen', False):
            sys.path.insert(0, str(Path(sys._MEIPASS) / "src"))
        else:
            sys.path.insert(0, str(Path(__file__).parent / "src"))
        
        from selenium import webdriver
        from selenium.webdriver.chrome.service import Service
        
        # Find ChromeDriver
        chromedriver_path = None
        possible_paths = [
            "chromedriver.exe",
            "src/chromedriver.exe",
            str(Path(__file__).parent / "chromedriver.exe"),
            str(Path(__file__).parent / "src" / "chromedriver.exe"),
        ]
        
        if getattr(sys, 'frozen', False):
            possible_paths.extend([
                str(Path(sys._MEIPASS) / "chromedriver.exe"),
                str(Path(sys._MEIPASS) / "src" / "chromedriver.exe"),
                str(Path(sys.executable).parent / "chromedriver.exe"),
            ])
        
        for path in possible_paths:
            if os.path.exists(path):
                chromedriver_path = path
                break
        
        if not chromedriver_path:
            print("❌ ChromeDriver not found - cannot test Selenium")
            return
        
        print(f"Using ChromeDriver at: {chromedriver_path}")
        
        # Set up Chrome options
        options = webdriver.ChromeOptions()
        options.add_argument("--headless")  # Run in headless mode for testing
        options.add_argument("--disable-gpu")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        
        # Create service
        service = Service(chromedriver_path)
        
        print("Starting Chrome browser...")
        driver = webdriver.Chrome(service=service, options=options)
        
        print("✅ Chrome browser started successfully")
        
        # Test navigation
        print("Testing navigation to Google...")
        driver.get("https://www.google.com")
        title = driver.title
        print(f"✅ Successfully navigated to Google (title: {title})")
        
        # Test oemsecrets.com
        print("Testing navigation to oemsecrets.com...")
        driver.get("https://www.oemsecrets.com")
        title = driver.title
        print(f"✅ Successfully navigated to oemsecrets.com (title: {title})")
        
        # Cleanup
        driver.quit()
        print("✅ Browser closed successfully")
        
        print("\n🎉 Selenium test completed successfully!")
        
    except Exception as e:
        print(f"❌ Selenium test failed: {e}")
        import traceback
        print(f"Full error: {traceback.format_exc()}")

def main():
    """Run all diagnostic checks."""
    print("BoMination ChromeDriver Diagnostic Tool")
    print("This tool helps identify why ChromeDriver might not be working.")
    print()
    
    check_system_info()
    check_chrome_installation()
    check_chromedriver_locations()
    check_network_connectivity()
    check_antivirus_interference()
    test_minimal_selenium()
    
    print("="*60)
    print("DIAGNOSTIC COMPLETE")
    print("="*60)
    print("If all checks pass but the application still doesn't work:")
    print("1. Try running the application as administrator")
    print("2. Check Windows Defender or antivirus exclusions")
    print("3. Try running from a different network")
    print("4. Contact support with the results of this diagnostic")
    print()
    
    input("Press Enter to exit...")

if __name__ == "__main__":
    main()
