import pandas as pd
import time
import os
import shutil
import sys
import traceback
from pathlib import Path
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import WebDriverException, TimeoutException, NoSuchElementException
from pipeline.validation_utils import handle_common_errors, check_chromedriver_availability, generate_output_path
import logging

# Suppress verbose Selenium logging
logging.getLogger('selenium').setLevel(logging.WARNING)
logging.getLogger('urllib3').setLevel(logging.WARNING)
logging.getLogger('requests').setLevel(logging.WARNING)
logging.getLogger('selenium.webdriver.remote.remote_connection').setLevel(logging.WARNING)

# Support both script and PyInstaller .exe paths for ChromeDriver
def find_chromedriver_path():
    """Find ChromeDriver in various possible locations."""
    if getattr(sys, 'frozen', False):
        # PyInstaller paths - check multiple locations
        possible_paths = [
            Path(sys._MEIPASS) / "chromedriver.exe",  # Root of extracted files
            Path(sys._MEIPASS) / "src" / "chromedriver.exe",  # src subdirectory
            Path(sys.executable).parent / "chromedriver.exe",  # Next to .exe
        ]
    else:
        # Development paths
        possible_paths = [
            Path(__file__).parent / "chromedriver.exe",  # Same directory as script
            Path(__file__).parent.parent / "chromedriver.exe",  # Parent directory (src/)
        ]
    
    print("🔍 Searching for ChromeDriver in the following locations:")
    for path in possible_paths:
        exists = path.exists()
        print(f"  {path} -> {'✓ Found' if exists else '✗ Not found'}")
        if exists:
            return str(path)
    
    # If not found, return the most likely path for error reporting
    return str(possible_paths[0]) if possible_paths else "chromedriver.exe"

CHROMEDRIVER_PATH = find_chromedriver_path()

def setup_browser():
    """Setup Chrome browser with error handling."""
    try:
        # Debug: Print ChromeDriver path being used
        print(f"🔍 Using ChromeDriver at: {CHROMEDRIVER_PATH}")
        print(f"🔍 ChromeDriver exists: {os.path.exists(CHROMEDRIVER_PATH)}")
        
        if getattr(sys, 'frozen', False):
            print(f"🔍 Running in PyInstaller mode")
            print(f"🔍 sys._MEIPASS: {sys._MEIPASS}")
            print(f"🔍 sys.executable: {sys.executable}")
        else:
            print(f"🔍 Running in development mode")
            print(f"🔍 Script directory: {Path(__file__).parent}")
        
        # Check ChromeDriver availability first
        chrome_available, chrome_version, chrome_error = check_chromedriver_availability()
        if not chrome_available:
            raise Exception(f"ChromeDriver setup failed: {chrome_error}")
        
        print(f"Using ChromeDriver: {chrome_version}")
        
        options = webdriver.ChromeOptions()
        # Start with minimal options for better compatibility
        options.add_argument("--disable-gpu")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--disable-extensions")
        
        # Reduce Chrome logging output
        options.add_argument("--log-level=3")
        options.add_argument("--silent")
        
        # Only add headless mode if specified (comment out for visible browser)
        # options.add_argument("--headless")
        
        # Create service with ChromeDriver path
        service = Service(CHROMEDRIVER_PATH)
        
        print("Starting Chrome browser...")
        print(f"ChromeDriver path: {CHROMEDRIVER_PATH}")
        driver = webdriver.Chrome(service=service, options=options)
        print("✅ Browser started successfully")
        
        # Set a reasonable page load timeout
        driver.set_page_load_timeout(30)
        
        return driver
        
    except WebDriverException as e:
        error_msg = handle_common_errors(str(e), WebDriverException)
        print(error_msg)
        raise Exception(error_msg)
    except Exception as e:
        error_msg = handle_common_errors(str(e))
        print(error_msg)
        raise

def upload_file_to_bomtool(driver, file_path):
    """Upload file to BoM tool with enhanced error handling."""
    try:
        print("🌐 Navigating to oemsecrets.com...")
        driver.get("https://www.oemsecrets.com")

        # Step 1: Go to BoM Tool
        print("🔗 Looking for BoM Tool link...")
        bom_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.LINK_TEXT, "BoM Tool"))
        )
        bom_button.click()
        print("✅ Clicked BoM Tool link")

        # Step 2: Upload File
        print("📤 Uploading file...")
        file_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input.filepond--browser"))
        )
        print(f"📄 Uploading: {file_path}")
        file_input.send_keys(file_path)

        # Step 3: Click "Import File"
        print("⏳ Clicking Import File button...")
        import_button = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "button.submit-selection"))
        )
        import_button.click()
        print("✅ File import initiated")

        # Step 4: Wait for processing to complete (with increased timeout)
        print("⏳ Waiting for file processing...")
        try:
            WebDriverWait(driver, 60).until(
                EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), 'Export') and contains(@class, 'dropdown-toggle')]"))
            )
            print("✅ Export dropdown is now visible - processing complete")
        except TimeoutException:
            print("⚠️ Export dropdown not visible after 60 seconds - checking if processing is still ongoing...")
            # Sometimes the processing takes longer, let's give it more time
            time.sleep(10)
            try:
                export_dropdown = driver.find_element(By.XPATH, "//a[contains(text(), 'Export') and contains(@class, 'dropdown-toggle')]")
                if export_dropdown.is_displayed():
                    print("✅ Export dropdown found after additional wait")
                else:
                    raise Exception("Export dropdown not found - file processing may have failed")
            except:
                raise Exception("File processing appears to have failed - export option not available")
        
        # ⏳ Wait for pricing data to load dynamically based on number of parts
        # Count the number of parts in the input file to determine wait time
        try:
            df = pd.read_excel(file_path)
            part_count = len(df)
            # Wait 1.2 seconds per part, with minimum 5 seconds and maximum 180 seconds
            wait_time = max(5, min(180, int(part_count * 1.2)))
            print(f"⏳ Found {part_count} parts in BOM - waiting {wait_time} seconds for pricing data to load...")
            time.sleep(wait_time)
        except Exception as e:
            # Fallback to 15 seconds if we can't read the file for some reason
            print(f"⚠️ Could not count parts ({e}), using default 15 second wait...")
            time.sleep(15)

        # Step 5: Click Export dropdown using JS (to ensure it actually opens)
        print("📥 Opening export dropdown...")
        export_dropdown = WebDriverWait(driver, 40).until(
            EC.presence_of_element_located((By.XPATH, "//a[contains(text(), 'Export') and contains(@class, 'dropdown-toggle')]"))
        )
        driver.execute_script("arguments[0].click();", export_dropdown)
        time.sleep(2)  # give time for dropdown to expand

        # Step 6: Wait for and select ".xlsx" radio option
        print("📊 Selecting XLSX format...")
        xlsx_radio = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='radio'][value='xlsx']"))
        )
        driver.execute_script("arguments[0].click();", xlsx_radio)  # JS click to handle visibility issues

        # Step 7: Click Download button
        print("⬇️ Starting download...")
        download_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "button.btn-tertiary"))
        )
        download_button.click()

        print("⏳ Waiting for file to download...")

        # Step 8: Wait for download and move it
        download_dir = os.path.join(os.path.expanduser("~"), "Downloads")
        timeout = 30
        filename = None
        
        for i in range(timeout):
            for file in os.listdir(download_dir):
                if file.endswith(".xlsx") and "bom" in file.lower():
                    filename = file
                    break
            if filename:
                break
            if i % 5 == 0:  # Print progress every 5 seconds
                print(f"⏳ Still waiting for download... ({i}/{timeout} seconds)")
            time.sleep(1)

        if not filename:
            raise Exception("Download failed or timed out after 30 seconds. "
                          "Please check your internet connection and try again.")

        src_path = os.path.join(download_dir, filename)
        
        # Save to the same directory as the input file (PDF directory)
        input_path_obj = Path(file_path)
        # Remove "_merged" suffix from stem to avoid double naming
        base_name = input_path_obj.stem.replace("_merged", "")
        dest_path = input_path_obj.parent / f"{base_name}_merged_with_prices.xlsx"
        
        print(f"📁 Moving price data to: {dest_path}")
        shutil.move(src_path, dest_path)

        print(f"✅ Price data saved to: {dest_path}")
        return dest_path
        
    except TimeoutException as e:
        error_msg = (
            "⏰ Timeout Error: The web page took too long to respond.\n\n"
            "Possible causes:\n"
            "• Slow internet connection\n"
            "• Server is busy or temporarily unavailable\n"
            "• Large file taking longer to process\n\n"
            "Solutions:\n"
            "• Check your internet connection\n"
            "• Try again in a few minutes\n"
            "• Use a smaller file if possible"
        )
        print(error_msg)
        raise Exception(error_msg)
        
    except NoSuchElementException as e:
        error_msg = (
            "🔍 Element Not Found: Could not find expected element on the webpage.\n\n"
            "Possible causes:\n"
            "• Website layout has changed\n"
            "• Page didn't load properly\n"
            "• Browser compatibility issue\n\n"
            "Solutions:\n"
            "• Try refreshing and running again\n"
            "• Check if the website is accessible in your browser\n"
            "• Update ChromeDriver if necessary"
        )
        print(error_msg)
        raise Exception(error_msg)
        
    except Exception as e:
        error_msg = handle_common_errors(str(e))
        print(error_msg)
        raise

def main():
    """Main function with comprehensive error handling."""
    input_path = os.environ.get("BOM_EXCEL_PATH")
    
    if not input_path:
        print("❌ Error: BOM_EXCEL_PATH environment variable not set")
        return
    
    if not os.path.exists(input_path):
        print(f"❌ Error: Input file does not exist: {input_path}")
        return
    
    print(f"📄 Processing file: {input_path}")
    
    driver = None
    try:
        driver = setup_browser()
        print("🔍 Starting price lookup process...")
        
        output_path = upload_file_to_bomtool(driver, os.path.abspath(input_path))
        
        if output_path:
            print(f"\n✅ Price lookup completed successfully!")
            print(f"📄 Output saved to: {output_path}")
        else:
            print("\n❌ Price lookup failed - no output file generated")
            
    except WebDriverException as e:
        error_msg = handle_common_errors(str(e), WebDriverException)
        print(f"\n❌ Browser automation error:\n{error_msg}")
    except TimeoutException as e:
        print(f"\n❌ Timeout error: The operation took too long to complete.\n"
              f"This might be due to slow internet connection or server issues.\n"
              f"Original error: {str(e)}")
    except Exception as e:
        print(f"❌ Error in price lookup: {e}")
        print(f"Traceback: {traceback.format_exc()}")
        
        # More detailed error information for common issues
        error_str = str(e).lower()
        if "certificate" in error_str or "ssl" in error_str or "connection is not private" in error_str:
            raise Exception(f"SSL/Certificate Error: The website has certificate issues that prevent automated access. "
                          f"This may require manual intervention or website policy changes. Original error: {e}")
        elif "timeout" in error_str:
            raise Exception(f"Network Timeout: The website is not responding. Check your internet connection or try again later. Original error: {e}")
        elif "webdriver" in error_str or "chromedriver" in error_str:
            raise Exception(f"Browser Driver Error: Problem with Chrome/ChromeDriver setup. Original error: {e}")
        else:
            raise Exception(f"Price Lookup Failed: {e}")
    
    finally:
        if driver:
            try:
                driver.quit()
                print("🔒 Browser closed successfully")
            except Exception as e:
                print(f"⚠️ Warning: Error closing browser: {str(e)}")


if __name__ == "__main__":
    main()