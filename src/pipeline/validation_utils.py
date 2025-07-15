"""
Input validation and error handling utilities for BoMination pipeline.
"""
import re
import os
import glob
import subprocess
import webbrowser
import shutil
import winreg
import sys
from pathlib import Path
from selenium.common.exceptions import WebDriverException


def validate_page_range(page_range_str):
    """
    Validate page range format and return parsed ranges.
    
    Args:
        page_range_str (str): Page range string (e.g., "1-3", "5", "2,4,6", "1-3,5,7-9")
    
    Returns:
        tuple: (is_valid, error_message, parsed_ranges)
    """
    if not page_range_str or not page_range_str.strip():
        return False, "Page range cannot be empty.", None
    
    page_range_str = page_range_str.strip()
    
    # Pattern to match valid page ranges: numbers, commas, hyphens, and spaces
    pattern = r'^[\d\s,\-]+$'
    if not re.match(pattern, page_range_str):
        return False, "Invalid characters in page range. Use only numbers, commas, and hyphens (e.g., '1-3,5,7-9').", None
    
    # Split by commas and validate each part
    parts = [part.strip() for part in page_range_str.split(',')]
    parsed_ranges = []
    
    for part in parts:
        if not part:
            continue
            
        if '-' in part:
            # Range like "1-3"
            try:
                start, end = part.split('-', 1)
                start_num = int(start.strip())
                end_num = int(end.strip())
                
                if start_num <= 0 or end_num <= 0:
                    return False, f"Page numbers must be positive integers. Found: {part}", None
                
                if start_num > end_num:
                    return False, f"Invalid range '{part}': start page cannot be greater than end page.", None
                
                parsed_ranges.append((start_num, end_num))
                
            except ValueError:
                return False, f"Invalid range format: '{part}'. Use format like '1-3'.", None
        else:
            # Single page like "5"
            try:
                page_num = int(part)
                if page_num <= 0:
                    return False, f"Page numbers must be positive integers. Found: {part}", None
                parsed_ranges.append(page_num)
            except ValueError:
                return False, f"Invalid page number: '{part}'. Must be a positive integer.", None
    
    if not parsed_ranges:
        return False, "No valid page numbers found in the range.", None
    
    return True, None, parsed_ranges


def validate_pdf_file(pdf_path):
    """
    Validate PDF file exists and is accessible.
    
    Args:
        pdf_path (str): Path to PDF file
    
    Returns:
        tuple: (is_valid, error_message)
    """
    if not pdf_path or not pdf_path.strip():
        return False, "PDF file path cannot be empty."
    
    pdf_path = pdf_path.strip()
    
    if not os.path.exists(pdf_path):
        return False, f"PDF file does not exist: {pdf_path}"
    
    if not pdf_path.lower().endswith('.pdf'):
        return False, f"File must be a PDF. Found: {os.path.splitext(pdf_path)[1]}"
    
    try:
        # Check if file is readable
        with open(pdf_path, 'rb') as f:
            f.read(4)  # Just read first 4 bytes to check accessibility
    except PermissionError:
        return False, f"Permission denied: Cannot access PDF file at {pdf_path}"
    except Exception as e:
        return False, f"Error accessing PDF file: {str(e)}"
    
    return True, None


def check_java_installation():
    """
    Check if Java is installed and accessible.
    Enhanced for PyInstaller compatibility with comprehensive detection methods.
    
    Returns:
        tuple: (is_installed, version_info, error_message)
    """
    debug_log = []
    debug_log.append("=== Java Detection Debug Log ===")
    
    # Method 1: Try the standard 'java' command first
    debug_log.append("Method 1: Trying 'java' command...")
    try:
        result = subprocess.run(
            ['java', '-version'], 
            capture_output=True, 
            text=True, 
            timeout=10,
            shell=True  # Use shell for better PATH resolution
        )
        
        debug_log.append(f"Java command result: returncode={result.returncode}")
        if result.returncode == 0:
            version_output = result.stderr or result.stdout
            debug_log.append(f"Java found via command: {version_output.strip()}")
            # Write debug log before returning
            try:
                with open("java_debug.log", "w") as f:
                    f.write("\n".join(debug_log))
            except:
                pass
            return True, version_output.strip(), None
        else:
            debug_log.append(f"Java command failed: stderr={result.stderr}, stdout={result.stdout}")
            
    except Exception as e:
        debug_log.append(f"Java command exception: {str(e)}")
        pass  # Continue to try other methods
    
    # Method 2: Check using shutil.which for PATH resolution
    debug_log.append("Method 2: Trying shutil.which...")
    try:
        java_path = shutil.which('java')
        debug_log.append(f"shutil.which result: {java_path}")
        if java_path:
            result = subprocess.run(
                [java_path, '-version'], 
                capture_output=True, 
                text=True, 
                timeout=10
            )
            debug_log.append(f"shutil.which java result: returncode={result.returncode}")
            if result.returncode == 0:
                version_output = result.stderr or result.stdout
                debug_log.append(f"Java found via shutil.which: {version_output.strip()}")
                # Write debug log before returning
                try:
                    with open("java_debug.log", "w") as f:
                        f.write("\n".join(debug_log))
                except:
                    pass
                return True, f"Found via shutil.which: {version_output.strip()}", None
    except Exception as e:
        debug_log.append(f"shutil.which exception: {str(e)}")
        pass
    
    # Method 3: Check Windows Registry for Java installations
    debug_log.append("Method 3: Checking Windows Registry...")
    try:
        def check_registry_key(hive, key_path):
            try:
                with winreg.OpenKey(hive, key_path) as key:
                    java_home, _ = winreg.QueryValueEx(key, "JavaHome")
                    java_exe = os.path.join(java_home, "bin", "java.exe")
                    if os.path.exists(java_exe):
                        return java_exe
            except (FileNotFoundError, OSError):
                pass
            return None
        
        # Check common registry locations
        registry_paths = [
            (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\JavaSoft\Java Runtime Environment"),
            (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\JavaSoft\Java Development Kit"),
            (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\JavaSoft\Java Runtime Environment"),
            (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\JavaSoft\Java Development Kit"),
            (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Eclipse Adoptium\JRE"),
            (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Eclipse Adoptium\JDK"),
        ]
        
        for hive, base_path in registry_paths:
            debug_log.append(f"Checking registry path: {base_path}")
            try:
                with winreg.OpenKey(hive, base_path) as base_key:
                    # Get current version
                    try:
                        current_version, _ = winreg.QueryValueEx(base_key, "CurrentVersion")
                        debug_log.append(f"Found current version: {current_version}")
                        version_path = f"{base_path}\\{current_version}"
                        java_exe = check_registry_key(hive, version_path)
                        if java_exe:
                            debug_log.append(f"Found Java exe via registry: {java_exe}")
                            result = subprocess.run(
                                [java_exe, '-version'], 
                                capture_output=True, 
                                text=True, 
                                timeout=10
                            )
                            if result.returncode == 0:
                                version_output = result.stderr or result.stdout
                                debug_log.append(f"Registry Java works: {version_output.strip()}")
                                # Write debug log before returning
                                try:
                                    with open("java_debug.log", "w") as f:
                                        f.write("\n".join(debug_log))
                                except:
                                    pass
                                return True, f"Found via registry: {version_output.strip()}", None
                    except FileNotFoundError:
                        debug_log.append("No CurrentVersion, enumerating subkeys...")
                        # If no CurrentVersion, enumerate all subkeys
                        i = 0
                        while True:
                            try:
                                version = winreg.EnumKey(base_key, i)
                                debug_log.append(f"Found version subkey: {version}")
                                version_path = f"{base_path}\\{version}"
                                java_exe = check_registry_key(hive, version_path)
                                if java_exe:
                                    debug_log.append(f"Found Java exe in version {version}: {java_exe}")
                                    result = subprocess.run(
                                        [java_exe, '-version'], 
                                        capture_output=True, 
                                        text=True, 
                                        timeout=10
                                    )
                                    if result.returncode == 0:
                                        version_output = result.stderr or result.stdout
                                        debug_log.append(f"Registry Java (version {version}) works: {version_output.strip()}")
                                        # Write debug log before returning
                                        try:
                                            with open("java_debug.log", "w") as f:
                                                f.write("\n".join(debug_log))
                                        except:
                                            pass
                                        return True, f"Found via registry: {version_output.strip()}", None
                                i += 1
                            except OSError:
                                break
            except (FileNotFoundError, OSError) as e:
                debug_log.append(f"Registry path not found: {str(e)}")
                continue
                
    except Exception as e:
        debug_log.append(f"Registry check exception: {str(e)}")
        pass  # Registry access might fail
    
    # Method 4: Check common installation paths with glob
    debug_log.append("Method 4: Checking common installation paths...")
    common_java_paths = [
        r"C:\Program Files\Java\jre*\bin\java.exe",
        r"C:\Program Files\Java\jdk*\bin\java.exe", 
        r"C:\Program Files (x86)\Java\jre*\bin\java.exe",
        r"C:\Program Files (x86)\Java\jdk*\bin\java.exe",
        r"C:\Program Files\Eclipse Adoptium\jre*\bin\java.exe",
        r"C:\Program Files\Eclipse Adoptium\jdk*\bin\java.exe",
        r"C:\Program Files\Microsoft\jdk*\bin\java.exe",
        r"C:\Program Files\Amazon Corretto\jdk*\bin\java.exe",
        r"C:\Program Files\Zulu\zulu*\bin\java.exe",
    ]
    
    for path_pattern in common_java_paths:
        debug_log.append(f"Checking path pattern: {path_pattern}")
        try:
            # Use glob to handle wildcards in paths
            matching_paths = glob.glob(path_pattern)
            debug_log.append(f"Glob matches: {matching_paths}")
            for java_path in matching_paths:
                if os.path.exists(java_path):
                    debug_log.append(f"Found existing Java path: {java_path}")
                    try:
                        result = subprocess.run(
                            [java_path, '-version'], 
                            capture_output=True, 
                            text=True, 
                            timeout=10
                        )
                        debug_log.append(f"Java path test result: returncode={result.returncode}")
                        if result.returncode == 0:
                            version_output = result.stderr or result.stdout
                            debug_log.append(f"Java path works: {version_output.strip()}")
                            # Write debug log before returning
                            try:
                                with open("java_debug.log", "w") as f:
                                    f.write("\n".join(debug_log))
                            except:
                                pass
                            return True, f"Found at {java_path}: {version_output.strip()}", None
                    except Exception as e:
                        debug_log.append(f"Java path test exception: {str(e)}")
                        continue
        except Exception as e:
            debug_log.append(f"Path pattern check exception: {str(e)}")
            continue
    
    # Method 5: Check if JAVA_HOME is set
    debug_log.append("Method 5: Checking JAVA_HOME...")
    java_home = os.environ.get('JAVA_HOME')
    debug_log.append(f"JAVA_HOME value: {java_home}")
    if java_home:
        java_exe = os.path.join(java_home, 'bin', 'java.exe')
        debug_log.append(f"JAVA_HOME java.exe path: {java_exe}")
        if os.path.exists(java_exe):
            debug_log.append("JAVA_HOME java.exe exists")
            try:
                result = subprocess.run(
                    [java_exe, '-version'], 
                    capture_output=True, 
                    text=True, 
                    timeout=10
                )
                debug_log.append(f"JAVA_HOME test result: returncode={result.returncode}")
                if result.returncode == 0:
                    version_output = result.stderr or result.stdout
                    debug_log.append(f"JAVA_HOME Java works: {version_output.strip()}")
                    # Write debug log before returning
                    try:
                        with open("java_debug.log", "w") as f:
                            f.write("\n".join(debug_log))
                    except:
                        pass
                    return True, f"Found via JAVA_HOME: {version_output.strip()}", None
            except Exception as e:
                debug_log.append(f"JAVA_HOME test exception: {str(e)}")
                pass
        else:
            debug_log.append("JAVA_HOME java.exe does not exist")
    
    debug_log.append("=== No Java found by any method ===")
    # Write debug log before returning failure
    try:
        with open("java_debug.log", "w") as f:
            f.write("\n".join(debug_log))
    except:
        pass
    
    return False, None, "Java not found. Please install Java or ensure it's in your system PATH to use PDF table extraction."


def check_chromedriver_availability():
    """
    Check if ChromeDriver is available and accessible.
    
    Returns:
        tuple: (is_available, version_info, error_message)
    """
    # Check in multiple possible locations for both development and PyInstaller
    possible_paths = [
        "chromedriver.exe",  # Current directory
        "src/chromedriver.exe",  # src subdirectory
        str(Path(__file__).parent / "chromedriver.exe"),  # Same as script
        str(Path(__file__).parent.parent / "chromedriver.exe"),  # Parent directory (src/)
    ]
    
    # Add PyInstaller-specific paths
    if getattr(sys, 'frozen', False):
        possible_paths.extend([
            str(Path(sys._MEIPASS) / "chromedriver.exe"),  # Root of extracted files
            str(Path(sys._MEIPASS) / "src" / "chromedriver.exe"),  # src subdirectory
            str(Path(sys.executable).parent / "chromedriver.exe"),  # Next to .exe
        ])
    
    print("🔍 Checking ChromeDriver availability...")
    for chromedriver_path in possible_paths:
        exists = os.path.exists(chromedriver_path)
        print(f"  {chromedriver_path} -> {'✓ Found' if exists else '✗ Not found'}")
        
        if exists:
            try:
                result = subprocess.run(
                    [chromedriver_path, '--version'], 
                    capture_output=True, 
                    text=True, 
                    timeout=10
                )
                
                if result.returncode == 0:
                    version_output = result.stdout.strip()
                    print(f"✓ ChromeDriver version: {version_output}")
                    return True, version_output, None
                else:
                    print(f"✗ ChromeDriver at {chromedriver_path} returned error code {result.returncode}")
                    continue
                    
            except subprocess.TimeoutExpired:
                print(f"✗ ChromeDriver at {chromedriver_path} timed out during version check")
                return False, None, f"ChromeDriver at {chromedriver_path} timed out during version check."
            except Exception as e:
                print(f"✗ ChromeDriver at {chromedriver_path} error: {e}")
                continue
    
    return False, None, "ChromeDriver not found. Please ensure chromedriver.exe is in the application directory or src folder."


def handle_common_errors(error_message, error_type=None):
    """
    Handle common errors with helpful user guidance.
    
    Args:
        error_message (str): The error message
        error_type (str): Type of error for specific handling
    
    Returns:
        str: User-friendly error message with guidance
    """
    error_lower = error_message.lower()
    
    # Java-related errors
    if 'java' in error_lower and ('not found' in error_lower or 'command not found' in error_lower):
        return (
            "❌ Java Not Found\n\n"
            "The PDF table extraction requires Java to be installed.\n\n"
            "Solutions:\n"
            "1. Install Java from: https://www.java.com/download/\n"
            "2. Restart your computer after installation\n"
            "3. Verify installation by opening Command Prompt and typing 'java -version'\n\n"
            "Would you like to open the Java download page?"
        )
    
    # ChromeDriver related errors
    if 'chromedriver' in error_lower or 'chrome' in error_lower:
        if 'not found' in error_lower or 'no such file' in error_lower:
            return (
                "❌ ChromeDriver Not Found\n\n"
                "The price lookup feature requires ChromeDriver.\n\n"
                "Solutions:\n"
                "1. Ensure chromedriver.exe is in the application folder\n"
                "2. Download ChromeDriver from: https://chromedriver.chromium.org/\n"
                "3. Make sure ChromeDriver version matches your Chrome browser version\n\n"
                "Original error: " + error_message
            )
        elif 'version' in error_lower or 'compatibility' in error_lower:
            return (
                "❌ ChromeDriver Version Mismatch\n\n"
                "ChromeDriver version doesn't match your Chrome browser.\n\n"
                "Solutions:\n"
                "1. Check your Chrome version (Chrome menu > Help > About Google Chrome)\n"
                "2. Download matching ChromeDriver from: https://chromedriver.chromium.org/\n"
                "3. Replace the existing chromedriver.exe with the new version\n\n"
                "Original error: " + error_message
            )
    
    # Selenium WebDriver errors
    if 'webdriver' in error_lower or isinstance(error_type, WebDriverException):
        return (
            "❌ Browser Automation Error\n\n"
            "There was a problem with the browser automation.\n\n"
            "Solutions:\n"
            "1. Close all Chrome browser windows and try again\n"
            "2. Check that Chrome browser is installed and up to date\n"
            "3. Restart the application\n"
            "4. Check your internet connection\n\n"
            "Original error: " + error_message
        )
    
    # Permission errors
    if 'permission' in error_lower or 'access' in error_lower:
        return (
            "❌ File Access Error\n\n"
            "Cannot access the required file or folder.\n\n"
            "Solutions:\n"
            "1. Close the PDF file if it's open in another application\n"
            "2. Run the application as Administrator\n"
            "3. Check that the file is not read-only\n"
            "4. Ensure you have write permissions to the output folder\n\n"
            "Original error: " + error_message
        )
    
    # Network errors
    if 'network' in error_lower or 'connection' in error_lower or 'timeout' in error_lower:
        return (
            "❌ Network Connection Error\n\n"
            "Cannot connect to the required online services.\n\n"
            "Solutions:\n"
            "1. Check your internet connection\n"
            "2. Try again in a few minutes\n"
            "3. Check if your firewall or antivirus is blocking the application\n"
            "4. Verify that https://www.oemsecrets.com is accessible in your browser\n\n"
            "Original error: " + error_message
        )
    
    # Return the original error if no specific handling is available
    return f"❌ An error occurred:\n\n{error_message}\n\nPlease check the details above and try again."


def open_help_url(url):
    """
    Open a help URL in the default browser.
    
    Args:
        url (str): URL to open
    """
    try:
        webbrowser.open(url)
        return True
    except Exception:
        return False


def validate_extracted_tables(tables):
    """
    Validate that tables were successfully extracted.
    
    Args:
        tables (list): List of extracted tables
    
    Returns:
        tuple: (is_valid, warning_message)
    """
    if not tables or len(tables) == 0:
        return False, (
            "⚠️ No Tables Extracted\n\n"
            "No tables were found in the specified PDF pages.\n\n"
            "Suggestions:\n"
            "1. Verify the page range contains tables\n"
            "2. Check if the PDF has text-based tables (not images)\n"
            "3. Try a different page range\n"
            "4. Consider using OCR for image-based tables"
        )
    
    # Check if tables are empty or contain only headers
    non_empty_tables = []
    for i, table in enumerate(tables):
        if hasattr(table, 'shape') and table.shape[0] > 1:  # More than just header
            non_empty_tables.append(table)
        elif hasattr(table, '__len__') and len(table) > 0:
            non_empty_tables.append(table)
    
    if not non_empty_tables:
        return False, (
            "⚠️ Empty Tables Found\n\n"
            "Tables were detected but appear to be empty or contain only headers.\n\n"
            "Suggestions:\n"
            "1. Check if the correct pages were specified\n"
            "2. Verify the PDF contains actual data tables\n"
            "3. Try adjusting the extraction settings"
        )
    
    if len(non_empty_tables) < len(tables):
        return True, (
            f"⚠️ Partial Success\n\n"
            f"Found {len(tables)} tables, but {len(tables) - len(non_empty_tables)} appear to be empty.\n"
            f"Proceeding with {len(non_empty_tables)} non-empty tables."
        )
    
    return True, None


def generate_output_path(input_file_path, suffix, output_directory=None, extension=".xlsx"):
    """
    Generate output file path based on input file and optional output directory.
    
    Args:
        input_file_path (str): Path to the input file
        suffix (str): Suffix to add to the filename (e.g., "_merged", "_extracted")
        output_directory (str, optional): Directory to save the output file
        extension (str): File extension (default: ".xlsx")
    
    Returns:
        str: Generated output file path
    """
    # Get the base filename without extension
    input_path = Path(input_file_path)
    base_name = input_path.stem
    
    # Generate output filename
    output_filename = f"{base_name}{suffix}{extension}"
    
    # Use output directory if provided, otherwise use input file's directory
    if output_directory and output_directory.strip():
        output_dir = Path(output_directory)
        # Create directory if it doesn't exist
        output_dir.mkdir(parents=True, exist_ok=True)
        output_path = output_dir / output_filename
    else:
        # Save in the same directory as the input file
        output_path = input_path.parent / output_filename
    
    return str(output_path)


def validate_output_directory(output_directory):
    """
    Validate output directory exists and is writable.
    
    Args:
        output_directory (str): Path to output directory
    
    Returns:
        tuple: (is_valid, error_message)
    """
    if not output_directory or not output_directory.strip():
        return True, None  # Empty is valid (will use input file directory)
    
    try:
        output_dir = Path(output_directory)
        
        # Check if directory exists
        if not output_dir.exists():
            # Try to create it
            output_dir.mkdir(parents=True, exist_ok=True)
            return True, None
        
        # Check if it's actually a directory
        if not output_dir.is_dir():
            return False, f"Path exists but is not a directory: {output_directory}"
        
        # Check if we can write to it by creating a test file
        test_file = output_dir / ".bomination_write_test"
        try:
            test_file.touch()
            test_file.unlink()  # Delete the test file
            return True, None
        except PermissionError:
            return False, f"Permission denied: Cannot write to directory {output_directory}"
            
    except Exception as e:
        return False, f"Error accessing output directory: {str(e)}"