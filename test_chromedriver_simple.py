#!/usr/bin/env python3
"""
Simple ChromeDriver test to verify functionality.
This script tests ChromeDriver setup without running the full pipeline.
"""

import sys
import os
from pathlib import Path

# Add src directory to path for imports
sys.path.insert(0, str(Path(__file__).parent / "src"))

from validation_utils import check_chromedriver_availability
from lookup_price import find_chromedriver_path, setup_browser

def test_chromedriver():
    """Test ChromeDriver functionality step by step."""
    print("="*50)
    print("ChromeDriver Test")
    print("="*50)
    
    # Step 1: Check ChromeDriver availability
    print("\n1. Checking ChromeDriver availability...")
    chrome_available, chrome_version, chrome_error = check_chromedriver_availability()
    
    if not chrome_available:
        print(f"❌ ChromeDriver not available: {chrome_error}")
        return False
    
    print(f"✅ ChromeDriver available: {chrome_version}")
    
    # Step 2: Test ChromeDriver path resolution
    print("\n2. Testing ChromeDriver path resolution...")
    chromedriver_path = find_chromedriver_path()
    print(f"ChromeDriver path: {chromedriver_path}")
    
    if not os.path.exists(chromedriver_path):
        print(f"❌ ChromeDriver not found at: {chromedriver_path}")
        return False
    
    print("✅ ChromeDriver file exists")
    
    # Step 3: Test browser setup
    print("\n3. Testing browser setup...")
    try:
        driver = setup_browser()
        print("✅ Browser setup successful")
        
        # Step 4: Test basic navigation
        print("\n4. Testing basic navigation...")
        driver.get("https://www.google.com")
        print("✅ Navigation successful")
        
        # Step 5: Cleanup
        print("\n5. Cleaning up...")
        driver.quit()
        print("✅ Browser closed successfully")
        
        print("\n🎉 All tests passed! ChromeDriver is working correctly.")
        return True
        
    except Exception as e:
        print(f"❌ Browser setup failed: {e}")
        return False

if __name__ == "__main__":
    success = test_chromedriver()
    if success:
        print("\n✅ ChromeDriver test completed successfully!")
        sys.exit(0)
    else:
        print("\n❌ ChromeDriver test failed!")
        sys.exit(1)
