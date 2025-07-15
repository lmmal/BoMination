"""
OCR Preprocessing Module for BoMination

This module handles OCR preprocessing of PDF files to make them searchable
and compatible with table extraction tools like Tabula and Camelot.
"""

import os
import sys
import tempfile
import subprocess
from pathlib import Path
import logging
import shutil

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def check_ocrmypdf_installation():
    """
    Check if ocrmypdf is available and properly installed.
    Returns (success, version_info, error_message)
    """
    try:
        import ocrmypdf
        # Test if we can run ocrmypdf
        result = subprocess.run([sys.executable, '-m', 'ocrmypdf', '--version'], 
                              capture_output=True, text=True, timeout=10)
        if result.returncode == 0:
            version_info = result.stdout.strip()
            logger.info(f"✅ OCRmyPDF available: {version_info}")
            return True, version_info, None
        else:
            error_msg = f"OCRmyPDF command failed: {result.stderr}"
            logger.error(f"❌ {error_msg}")
            return False, None, error_msg
    except ImportError:
        error_msg = "OCRmyPDF not installed. Install with: pip install ocrmypdf"
        logger.error(f"❌ {error_msg}")
        return False, None, error_msg
    except subprocess.TimeoutExpired:
        error_msg = "OCRmyPDF check timed out"
        logger.error(f"❌ {error_msg}")
        return False, None, error_msg
    except Exception as e:
        error_msg = f"OCRmyPDF check failed: {str(e)}"
        logger.error(f"❌ {error_msg}")
        return False, None, error_msg

def check_tesseract_installation():
    """
    Check if Tesseract OCR is available (required by OCRmyPDF).
    Returns (success, version_info, error_message)
    """
    # List of common Tesseract installation paths on Windows
    common_paths = [
        'tesseract',  # From PATH
        r'C:\Program Files\Tesseract-OCR\tesseract.exe',
        r'C:\Program Files (x86)\Tesseract-OCR\tesseract.exe',
        r'C:\ProgramData\chocolatey\bin\tesseract.exe',
        r'C:\ProgramData\chocolatey\lib\tesseract\tools\tesseract.exe',
        r'C:\tools\tesseract\tesseract.exe',
    ]
    
    for tesseract_path in common_paths:
        try:
            # Try to run tesseract --version
            result = subprocess.run([tesseract_path, '--version'], 
                                  capture_output=True, text=True, timeout=10)
            if result.returncode == 0:
                version_info = result.stdout.split('\n')[0]  # First line has version
                logger.info(f"✅ Tesseract available: {version_info}")
                return True, version_info, None
            else:
                continue  # Try next path
        except FileNotFoundError:
            continue  # Try next path
        except subprocess.TimeoutExpired:
            continue  # Try next path
        except Exception:
            continue  # Try next path
    
    # If we get here, Tesseract was not found in any of the common locations
    error_msg = "Tesseract not found. Please install Tesseract OCR engine."
    logger.error(f"❌ {error_msg}")
    return False, None, error_msg

def is_pdf_searchable(pdf_path):
    """
    Check if a PDF already has searchable text.
    Returns True if the PDF contains searchable text, False otherwise.
    """
    try:
        import PyPDF2
        with open(pdf_path, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            
            # Check first few pages for text content
            text_found = False
            pages_to_check = min(3, len(pdf_reader.pages))
            
            for page_num in range(pages_to_check):
                page = pdf_reader.pages[page_num]
                text = page.extract_text().strip()
                if text and len(text) > 10:  # Meaningful text content
                    text_found = True
                    break
            
            if text_found:
                logger.info(f"✅ PDF appears to have searchable text")
                return True
            else:
                logger.info(f"⚠️ PDF appears to be image-based or has no searchable text")
                return False
                
    except ImportError:
        logger.warning("PyPDF2 not available for text detection, assuming OCR is needed")
        return False
    except Exception as e:
        logger.warning(f"Could not check if PDF is searchable: {e}")
        return False

def preprocess_pdf_with_ocr(pdf_path, output_path=None, force_ocr=False):
    """
    Preprocess a PDF file with OCR to make it searchable.
    
    Args:
        pdf_path (str): Path to the input PDF file
        output_path (str, optional): Path for the output OCR'd PDF. If None, uses temp file.
        force_ocr (bool): Force OCR even if PDF appears to have text
    
    Returns:
        tuple: (success, ocr_pdf_path, error_message)
    """
    pdf_path = Path(pdf_path)
    
    # Validate input
    if not pdf_path.exists():
        error_msg = f"PDF file not found: {pdf_path}"
        logger.error(f"❌ {error_msg}")
        return False, None, error_msg
    
    # Check if OCR is available
    ocr_available, ocr_version, ocr_error = check_ocrmypdf_installation()
    if not ocr_available:
        return False, None, ocr_error
    
    # Check if Tesseract is available
    tesseract_available, tesseract_version, tesseract_error = check_tesseract_installation()
    if not tesseract_available:
        return False, None, tesseract_error
    
    # Check if PDF already has searchable text (unless forced)
    if not force_ocr and is_pdf_searchable(pdf_path):
        logger.info("📝 PDF already appears to have searchable text, skipping OCR")
        return True, str(pdf_path), None
    
    # Setup output path
    if output_path is None:
        # Create a temporary file for OCR'd PDF
        temp_dir = tempfile.mkdtemp(prefix="bomination_ocr_")
        output_path = Path(temp_dir) / f"{pdf_path.stem}_ocr.pdf"
    else:
        output_path = Path(output_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)
    
    logger.info(f"🔍 Starting OCR preprocessing: {pdf_path.name}")
    logger.info(f"📄 Input: {pdf_path}")
    logger.info(f"📄 Output: {output_path}")
    
    try:
        # Import ocrmypdf here to avoid import issues if not available
        import ocrmypdf
        
        # OCR the PDF with optimized settings for table extraction
        ocrmypdf.ocr(
            input_file=str(pdf_path),
            output_file=str(output_path),
            deskew=True,          # Correct page rotation
            remove_background=False,  # Keep backgrounds for table detection
            optimize=1,           # Light optimization to keep structure
            force_ocr=force_ocr,  # Force OCR even if text detected
            skip_text=False,      # Don't skip existing text
            redo_ocr=False,       # Don't redo OCR if already done
            language='eng',       # English language for OCR
            oversample=600,       # Rasterize at 600 DPI instead of 400 for better table lines
            image_dpi=600,        # Ensure embedded images use 600 DPI
            progress_bar=False,   # Disable progress bar for cleaner output
            quiet=True           # Reduce verbose output
        )
        
        if output_path.exists():
            logger.info(f"✅ OCR preprocessing completed successfully")
            logger.info(f"📁 OCR'd PDF saved to: {output_path}")
            return True, str(output_path), None
        else:
            error_msg = "OCR completed but output file not found"
            logger.error(f"❌ {error_msg}")
            return False, None, error_msg
            
    except Exception as e:
        error_msg = f"OCR preprocessing failed: {str(e)}"
        logger.error(f"❌ {error_msg}")
        
        # Clean up partial output
        if output_path.exists():
            try:
                output_path.unlink()
            except:
                pass
        
        return False, None, error_msg

def cleanup_ocr_temp_files(ocr_pdf_path):
    """
    Clean up temporary OCR files.
    
    Args:
        ocr_pdf_path (str): Path to the OCR'd PDF file to clean up
    """
    if not ocr_pdf_path:
        return
        
    try:
        ocr_path = Path(ocr_pdf_path)
        
        # If it's in a temp directory we created, remove the entire directory
        if "bomination_ocr_" in str(ocr_path.parent):
            temp_dir = ocr_path.parent
            if temp_dir.exists():
                shutil.rmtree(temp_dir)
                logger.info(f"🧹 Cleaned up OCR temp directory: {temp_dir}")
        else:
            # Just remove the file
            if ocr_path.exists():
                ocr_path.unlink()
                logger.info(f"🧹 Cleaned up OCR file: {ocr_path}")
                
    except Exception as e:
        logger.warning(f"⚠️ Could not clean up OCR temp files: {e}")

def get_ocr_installation_instructions():
    """
    Get platform-specific installation instructions for OCR dependencies.
    """
    import platform
    
    system = platform.system().lower()
    
    if system == "windows":
        return """
To install OCR support on Windows:

1. Install OCRmyPDF:
   pip install ocrmypdf

2. Install Tesseract OCR engine:
   - Download from: https://github.com/UB-Mannheim/tesseract/wiki
   - Or use chocolatey: choco install tesseract
   - Or use conda: conda install -c conda-forge tesseract

3. Make sure Tesseract is in your PATH
   - Add the Tesseract installation directory to your system PATH
   - Default location: C:\\Program Files\\Tesseract-OCR\\

4. Restart your application after installation
"""
    elif system == "darwin":  # macOS
        return """
To install OCR support on macOS:

1. Install OCRmyPDF:
   pip install ocrmypdf

2. Install Tesseract OCR engine:
   brew install tesseract

3. Restart your application after installation
"""
    else:  # Linux
        return """
To install OCR support on Linux:

1. Install OCRmyPDF:
   pip install ocrmypdf

2. Install Tesseract OCR engine:
   # Ubuntu/Debian:
   sudo apt-get install tesseract-ocr
   
   # CentOS/RHEL:
   sudo yum install tesseract
   
   # Arch Linux:
   sudo pacman -S tesseract

3. Restart your application after installation
"""

# Test function for development
def test_ocr_functionality():
    """Test OCR functionality with a sample PDF."""
    print("Testing OCR functionality...")
    
    # Check installations
    print("\n1. Checking OCRmyPDF installation...")
    ocr_ok, ocr_version, ocr_error = check_ocrmypdf_installation()
    if not ocr_ok:
        print(f"❌ OCRmyPDF: {ocr_error}")
        print(get_ocr_installation_instructions())
        return False
    print(f"✅ OCRmyPDF: {ocr_version}")
    
    print("\n2. Checking Tesseract installation...")
    tesseract_ok, tesseract_version, tesseract_error = check_tesseract_installation()
    if not tesseract_ok:
        print(f"❌ Tesseract: {tesseract_error}")
        print(get_ocr_installation_instructions())
        return False
    print(f"✅ Tesseract: {tesseract_version}")
    
    print("\n✅ All OCR components are available!")
    return True

if __name__ == "__main__":
    # Run test when module is executed directly
    test_ocr_functionality()
