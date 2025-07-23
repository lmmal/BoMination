"""
Tabula-specific table extraction module for BoMination.
Handles table extraction from PDFs using the tabula-py library.
"""

import os
import sys
import logging
from pathlib import Path
import traceback
import pandas as pd

# Configure encoding for Windows compatibility
os.environ['PYTHONIOENCODING'] = 'utf-8:ignore'

# Add the src directory to Python path so we can import modules
src_dir = Path(__file__).parent.parent
if str(src_dir) not in sys.path:
    sys.path.insert(0, str(src_dir))

# Tabula-specific extraction functions
def configure_tabula_environment():
    """Configure e                           # Sa                # Save individual extracted tables first (preserves original table structure)
                print(f"💾 Saving individual extracted tables to: {extracted_path}")
                print(f"🔧 Using processed_tables for extracted file (count: {len(processed_tables)})")
                
                try:
                    extracted_success = save_tables_to_excel(processed_tables, str(extracted_path))
                    print(f"🔧 save_tables_to_excel returned: {extracted_success}")
                except Exception as save_error:
                    print(f"❌ save_tables_to_excel failed with exception: {save_error}")
                    import traceback
                    traceback.print_exc()
                    extracted_success = False
                
                if extracted_success:
                    print(f"✅ Individual tables saved: {extracted_path}")
                    print(f"   File exists: {extracted_path.exists()}")
                    if extracted_path.exists():
                        print(f"   File size: {extracted_path.stat().st_size} bytes")
                else:
                    print(f"❌ Failed to save individual tables to: {extracted_path}")
                    print(f"   extracted_success = {extracted_success}")idual extracted tables first (preserves original table structure)
                print(f"💾 Saving individual extracted tables to: {extracted_path}")
                print(f"🔧 Using processed_tables for extracted file (count: {len(processed_tables)})")
                
                try:
                    extracted_success = save_tables_to_excel(processed_tables, str(extracted_path))
                    print(f"🔧 save_tables_to_excel returned: {extracted_success}")
                except Exception as save_error:
                    print(f"❌ save_tables_to_excel failed with exception: {save_error}")
                    import traceback
                    traceback.print_exc()
                    extracted_success = False  # Save individual extracted tables first (preserves original table structure)
                print(f"💾 Saving individual extracted tables to: {extracted_path}")
                extracted_success = save_tables_to_excel(processed_tables, str(extracted_path))
                
                if extracted_success:
                    print(f"✅ Individual tables saved: {extracted_path}")
                else:
                    print(f"❌ Failed to save individual tables to: {extracted_path}")
                
                # Save merged tables (combines all tables into one)
                print(f"💾 Saving merged table to: {merged_path}")
                merged_success = merge_tables_and_export(processed_tables, str(merged_path))ariables for tabula to handle encoding issues."""
    original_java_options = os.environ.get('JAVA_TOOL_OPTIONS', '')
    original_lang = os.environ.get('LANG', '')
    original_lc_all = os.environ.get('LC_ALL', '')
    
    # Set environment to handle subprocess encoding issues
    os.environ['PYTHONIOENCODING'] = 'utf-8:ignore'
    os.environ['LANG'] = 'C.UTF-8'
    os.environ['LC_ALL'] = 'C.UTF-8'
    os.environ['JAVA_TOOL_OPTIONS'] = '-Dfile.encoding=UTF-8 -Duser.language=en -Duser.country=US -Djava.awt.headless=true'
    
    return original_java_options, original_lang, original_lc_all

def restore_tabula_environment(original_java_options, original_lang, original_lc_all):
    """Restore original environment variables after tabula execution."""
    # Restore original environment variables
    if original_java_options:
        os.environ['JAVA_TOOL_OPTIONS'] = original_java_options
    elif 'JAVA_TOOL_OPTIONS' in os.environ:
        del os.environ['JAVA_TOOL_OPTIONS']
        
    if original_lang:
        os.environ['LANG'] = original_lang
    elif 'LANG' in os.environ:
        del os.environ['LANG']
        
    if original_lc_all:
        os.environ['LC_ALL'] = original_lc_all
    elif 'LC_ALL' in os.environ:
        del os.environ['LC_ALL']
        
    if 'PYTHONIOENCODING' in os.environ:
        del os.environ['PYTHONIOENCODING']


def extract_tables_with_tabula_method_impl(pdf_path, pages):
    """
    Extract tables using Tabula-py with optimized settings.
    Works best with text-based PDFs.
    
    Args:
        pdf_path (str): Path to the PDF file
        pages (str): Page range to extract from (e.g., "1-3" or "all")
        
    Returns:
        list: List of extracted DataFrames
    """
    print("📊 Attempting table extraction with Tabula...")
    
    try:
        import tabula
        print(f"📊 Tabula version: {tabula.__version__}")
    except ImportError:
        print("❌ Tabula not available - install with: pip install tabula-py")
        return []
    
    # Configure environment
    original_java_options, original_lang, original_lc_all = configure_tabula_environment()
    
    try:
        tables = []
        
        # Method 1: Lattice detection (most reliable for structured tables)
        print("📊 Trying lattice detection...")
        try:
            lattice_tables = tabula.read_pdf(
                pdf_path,
                pages=pages,
                multiple_tables=True,
                lattice=True,
                guess=False,
                stream=False,
                java_options="-Dfile.encoding=UTF-8 -Duser.language=en -Duser.country=US -Djava.awt.headless=true",
                pandas_options={'header': None}
            )
            
            if lattice_tables:
                print(f"    ✅ Lattice method found {len(lattice_tables)} tables")
                for i, table in enumerate(lattice_tables):
                    if not table.empty and table.shape[0] >= 3 and table.shape[1] >= 3:
                        print(f"    Table {i+1}: {table.shape[0]}×{table.shape[1]} (lattice)")
                        tables.append(table)
                    else:
                        print(f"    Table {i+1}: {table.shape[0]}×{table.shape[1]} - too small (lattice)")
            else:
                print("    ❌ Lattice method found no tables")
                
        except Exception as e:
            print(f"    ❌ Lattice method failed: {e}")
        
        # Method 2: Stream detection (good for tables without clear borders)
        if not tables:
            print("📊 Trying stream detection...")
            try:
                stream_tables = tabula.read_pdf(
                    pdf_path,
                    pages=pages,
                    multiple_tables=True,
                    lattice=False,
                    stream=True,
                    guess=False,
                    java_options="-Dfile.encoding=UTF-8 -Duser.language=en -Duser.country=US -Djava.awt.headless=true",
                    pandas_options={'header': None}
                )
                
                if stream_tables:
                    print(f"    ✅ Stream method found {len(stream_tables)} tables")
                    for i, table in enumerate(stream_tables):
                        if not table.empty and table.shape[0] >= 3 and table.shape[1] >= 3:
                            print(f"    Table {i+1}: {table.shape[0]}×{table.shape[1]} (stream)")
                            tables.append(table)
                        else:
                            print(f"    Table {i+1}: {table.shape[0]}×{table.shape[1]} - too small (stream)")
                else:
                    print("    ❌ Stream method found no tables")
                    
            except Exception as e:
                print(f"    ❌ Stream method failed: {e}")
        
        # Method 3: Auto-detection with guess=True (most permissive)
        if not tables:
            print("📊 Trying auto-detection...")
            try:
                auto_tables = tabula.read_pdf(
                    pdf_path,
                    pages=pages,
                    multiple_tables=True,
                    lattice=True,
                    stream=False,
                    guess=True,
                    java_options="-Dfile.encoding=UTF-8 -Duser.language=en -Duser.country=US -Djava.awt.headless=true",
                    pandas_options={'header': None}
                )
                
                if auto_tables:
                    print(f"    ✅ Auto-detection found {len(auto_tables)} tables")
                    for i, table in enumerate(auto_tables):
                        if not table.empty and table.shape[0] >= 3 and table.shape[1] >= 3:
                            print(f"    Table {i+1}: {table.shape[0]}×{table.shape[1]} (auto)")
                            tables.append(table)
                        else:
                            print(f"    Table {i+1}: {table.shape[0]}×{table.shape[1]} - too small (auto)")
                else:
                    print("    ❌ Auto-detection found no tables")
                    
            except Exception as e:
                print(f"    ❌ Auto-detection failed: {e}")
        
        # Method 4: Last resort - very permissive settings
        if not tables:
            print("📊 Trying last resort extraction...")
            try:
                last_resort_tables = tabula.read_pdf(
                    pdf_path,
                    pages=pages,
                    multiple_tables=True,
                    lattice=True,
                    stream=True,
                    guess=True,
                    java_options="-Dfile.encoding=UTF-8 -Duser.language=en -Duser.country=US -Djava.awt.headless=true",
                    pandas_options={'header': None}
                )
                
                if last_resort_tables:
                    print(f"    ✅ Last resort found {len(last_resort_tables)} tables")
                    for i, table in enumerate(last_resort_tables):
                        if not table.empty and table.shape[0] >= 2 and table.shape[1] >= 2:  # Lower threshold
                            print(f"    Table {i+1}: {table.shape[0]}×{table.shape[1]} (last resort)")
                            tables.append(table)
                        else:
                            print(f"    Table {i+1}: {table.shape[0]}×{table.shape[1]} - too small (last resort)")
                else:
                    print("    ❌ Last resort found no tables")
                    
            except Exception as e:
                print(f"    ❌ Last resort failed: {e}")
        
        print(f"📊 Tabula extraction completed: {len(tables)} tables found")
        
        # Show summary of extracted tables
        if tables:
            print("📋 Tabula extraction summary:")
            for i, table in enumerate(tables):
                non_empty_cells = table.notna().sum().sum()
                total_cells = table.shape[0] * table.shape[1]
                density = non_empty_cells / total_cells if total_cells > 0 else 0
                print(f"    Table {i+1}: {table.shape[0]}×{table.shape[1]} (density: {density:.2f})")
                
                # Show sample content
                sample_text = ' '.join(table.fillna('').astype(str).values.flatten()[:10])
                print(f"    Sample: {sample_text[:100]}...")
        
        return tables
        
    except Exception as e:
        print(f"❌ Tabula extraction failed: {e}")
        print(f"🔍 Error type: {type(e).__name__}")
        traceback.print_exc()
        return []
        
    finally:
        # Restore original environment variables
        restore_tabula_environment(original_java_options, original_lang, original_lc_all)

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')
logger = logging.getLogger(__name__)

# Configure Tabula for PyInstaller
if getattr(sys, 'frozen', False):
    # Set Java options for PyInstaller
    os.environ['JAVA_TOOL_OPTIONS'] = '-Dfile.encoding=UTF-8 -Duser.language=en -Duser.country=US -Djava.awt.headless=true'

# Import customer formatting
from omni_cust.customer_formatters import apply_customer_formatter, get_available_customers

# Import validation utilities
from pipeline.validation_utils import validate_extracted_tables, handle_common_errors, generate_output_path

def check_roi_text_content(pdf_path, page_num, roi_area):
    """
    Check if there's extractable text content within the ROI area.
    
    Args:
        pdf_path (str): Path to the PDF file
        page_num (int): Page number (1-based)
        roi_area (list): ROI area as [top, left, bottom, right] in points
        
    Returns:
        bool: True if text content is found, False otherwise
    """
    try:
        import pdfplumber
        
        with pdfplumber.open(pdf_path) as pdf:
            if page_num > len(pdf.pages):
                print(f"❌ Page {page_num} not found in PDF")
                return False
            
            page = pdf.pages[page_num - 1]  # Convert to 0-based index
            
            # Convert ROI area from tabula format to pdfplumber format
            # Tabula: [top, left, bottom, right] in points from top-left
            # pdfplumber: uses (left, bottom, right, top) from bottom-left
            top, left, bottom, right = roi_area
            
            # Convert coordinates (pdfplumber uses bottom-left origin)
            page_height = page.height
            bbox = (left, page_height - bottom, right, page_height - top)
            
            # Extract text from the ROI area
            roi_text = page.crop(bbox).extract_text()
            
            if roi_text and roi_text.strip():
                print(f"✅ ROI area contains text: {len(roi_text.strip())} characters")
                return True
            else:
                print(f"❌ ROI area contains no extractable text")
                return False
                
    except ImportError:
        print("⚠️ pdfplumber not available - assuming text content exists")
        return True
    except Exception as e:
        print(f"❌ Error checking ROI text content: {e}")
        return True  # Default to assuming text exists

def extract_tables_with_roi_selection_tabula(pdf_path, pages, roi_callback=None):
    """
    Extract tables from PDF using manual ROI (Region of Interest) selection.
    
    ROI areas should be provided via environment variable BOM_ROI_AREAS as JSON.
    
    Args:
        pdf_path (str): Path to the PDF file
        pages (str): Page range to extract from
        roi_callback (callable): Optional callback function (deprecated, use environment)
    """
    print("\n📍 Starting ROI-based table extraction...")
    
    try:
        # Get ROI areas from environment variable
        roi_areas = None
        roi_areas_json = os.environ.get("BOM_ROI_AREAS")
        
        if roi_areas_json:
            import json
            try:
                roi_areas = json.loads(roi_areas_json)
                print(f"✅ ROI areas loaded from environment: {len(roi_areas)} areas")
            except json.JSONDecodeError as e:
                print(f"❌ Failed to parse ROI areas from environment: {e}")
                return []
        else:
            print("❌ No ROI areas found in environment variable BOM_ROI_AREAS")
            return []
        
        if not roi_areas:
            print("❌ No ROI areas available for extraction")
            return []
        
        print(f"✅ ROI areas loaded for {len(roi_areas)} pages")
        
        # Debug: Show what areas were loaded
        for page_num, area in roi_areas.items():
            print(f"📍 Page {page_num}: ROI area = {area}")
        
        # Extract tables using the provided areas
        all_tables = []
        ocr_pdf_path = None
        
        # Check if we're already working with an OCR-processed PDF
        is_ocr_processed = "_ocr" in Path(pdf_path).stem
        
        if is_ocr_processed:
            print(f"🔍 Detected OCR-processed PDF: {pdf_path}")
            print(f"⚠️ ROI coordinates may not be accurate due to OCR layout changes")
            working_pdf_path = pdf_path
            # Skip text content checking for OCR-processed PDFs since layout may have changed
        else:
            # Check if ROI areas contain text content
            print(f"\n🔍 Checking ROI areas for text content...")
            roi_needs_ocr = False
            
            for page_num, area in roi_areas.items():
                print(f"🔍 Checking page {page_num} ROI area for text content...")
                has_text = check_roi_text_content(pdf_path, int(page_num), area)
                if not has_text:
                    print(f"❌ Page {page_num} ROI area has no text - will need OCR")
                    roi_needs_ocr = True
                else:
                    print(f"✅ Page {page_num} ROI area has text - can use direct extraction")
            
            # If any ROI area needs OCR, process the entire PDF with OCR first
            working_pdf_path = pdf_path
            if roi_needs_ocr:
                print(f"\n🔄 ROI areas need OCR - preprocessing entire PDF...")
                
                try:
                    # Check if OCR preprocessor is available
                    try:
                        from pipeline.ocr_preprocessor import preprocess_pdf_with_ocr
                        print("🔧 OCR preprocessor imported successfully")
                    except ImportError as import_error:
                        print(f"❌ OCR preprocessor import failed: {import_error}")
                        print("⚠️ Continuing with original PDF despite no text in ROI")
                    else:
                        # Create OCR version of the PDF using proper path handling
                        pdf_path_obj = Path(pdf_path)
                        ocr_pdf_path = pdf_path_obj.parent / f"{pdf_path_obj.stem}_ocr.pdf"
                        
                        print(f"🔧 OCR input: {pdf_path}")
                        print(f"🔧 OCR output: {ocr_pdf_path}")
                        
                        success = preprocess_pdf_with_ocr(pdf_path, str(ocr_pdf_path))
                        
                        print(f"🔧 OCR success: {success}")
                        print(f"🔧 OCR file exists: {ocr_pdf_path.exists()}")
                        
                        if success and ocr_pdf_path.exists():
                            print(f"✅ OCR preprocessing successful: {ocr_pdf_path}")
                            working_pdf_path = str(ocr_pdf_path)
                        else:
                            print("❌ OCR preprocessing failed")
                            if not success:
                                print("❌ OCR function returned False")
                            if not ocr_pdf_path.exists():
                                print("❌ OCR output file was not created")
                            print("⚠️ Continuing with original PDF")
                            
                except Exception as e:
                    print(f"❌ OCR preprocessing error: {e}")
                    import traceback
                    print("📋 OCR error traceback:")
                    traceback.print_exc()
                    print("⚠️ Continuing with original PDF")
        
        # Now attempt extraction with the appropriate PDF
        print(f"\n🔄 Attempting extraction from: {working_pdf_path}")
        
        # First attempt: try extraction with current PDF (original or OCR'd)
        extraction_attempts = 0
        
        while extraction_attempts < 1 and not all_tables:  # Only one attempt needed now
            extraction_attempts += 1
            
            print(f"\n🔄 Attempt {extraction_attempts}: Extracting from {'OCR-processed' if working_pdf_path != pdf_path else 'original'} PDF...")
            
            for page_num, area in roi_areas.items():
                print(f"\n📊 Extracting table from page {page_num} using ROI area: {area}")
                
                try:
                    import tabula
                    
                    print(f"🔧 Tabula extraction debug:")
                    print(f"   Working PDF: {working_pdf_path}")
                    print(f"   PDF exists: {os.path.exists(working_pdf_path)}")
                    print(f"   Page: {page_num}")
                    print(f"   ROI area: {area}")
                    
                    # For OCR-processed PDFs, skip tabula ROI extraction since coordinates may not be accurate
                    # Let the system fall back to Camelot for better OCR-processed PDF handling
                    if is_ocr_processed:
                        print(f"   🔍 OCR-processed PDF detected - skipping tabula ROI extraction")
                        print(f"   ⚠️ ROI coordinates may not be accurate after OCR processing")
                        print(f"   🔄 Returning empty to trigger Camelot fallback...")
                        tables = []
                    else:
                        # Original logic for non-OCR processed PDFs
                        print(f"   🔄 Using original ROI extraction logic...")
                        
                        # First try with area restriction
                        tables = tabula.read_pdf(
                            working_pdf_path,
                            pages=[int(page_num)],
                            area=area,  # [top, left, bottom, right] in points
                            multiple_tables=False,
                            lattice=True,
                            guess=False,
                            java_options="-Dfile.encoding=UTF-8 -Duser.language=en -Duser.country=US -Djava.awt.headless=true",
                            pandas_options={'header': None}
                        )
                    
                    print(f"🔧 Tabula returned: {type(tables)}")
                    if tables:
                        print(f"🔧 Number of tables: {len(tables)}")
                        for i, table in enumerate(tables):
                            print(f"🔧 Table {i+1}: shape={table.shape}, empty={table.empty}")
                            if not table.empty and table.shape[0] >= 2 and table.shape[1] >= 2:
                                extraction_method = "OCR-skipped" if is_ocr_processed else "ROI"
                                print(f"    ✅ Extracted table {i+1}: {table.shape[0]}×{table.shape[1]} ({extraction_method})")
                                all_tables.append(table)
                            else:
                                print(f"    ❌ Table {i+1} too small: {table.shape[0]}×{table.shape[1]}")
                    else:
                        print(f"    ❌ No tables extracted from page {page_num}")
                        print(f"    ❌ Tabula returned: {tables}")
                        
                        # For OCR-processed PDFs, we expect this to fail so Camelot can handle it
                        if is_ocr_processed:
                            print(f"    ✅ Expected result for OCR-processed PDF - Camelot will handle this")
                        else:
                            # Only try fallback for non-OCR processed PDFs
                            print(f"    🔄 Trying without ROI area restriction...")
                            try:
                                fallback_tables = tabula.read_pdf(
                                    working_pdf_path,
                                    pages=[int(page_num)],
                                    multiple_tables=True,
                                    lattice=True,
                                    guess=False,
                                    java_options="-Dfile.encoding=UTF-8 -Duser.language=en -Duser.country=US -Djava.awt.headless=true",
                                    pandas_options={'header': None}
                                )
                                
                                if fallback_tables:
                                    print(f"    ✅ Fallback found {len(fallback_tables)} tables without ROI")
                                    for i, table in enumerate(fallback_tables):
                                        if not table.empty and table.shape[0] >= 2 and table.shape[1] >= 2:
                                            print(f"    ✅ Extracted fallback table {i+1}: {table.shape[0]}×{table.shape[1]} (no ROI)")
                                            all_tables.append(table)
                                        else:
                                            print(f"    ❌ Fallback table {i+1} too small: {table.shape[0]}×{table.shape[1]}")
                                else:
                                    print(f"    ❌ No fallback tables found either")
                            except Exception as fallback_error:
                                print(f"    ❌ Fallback extraction failed: {fallback_error}")
                    
                    # Note: Camelot fallback is now handled by extract_main.py orchestrator
                    # This module only handles tabula-specific extraction
                    
                except Exception as e:
                    print(f"❌ Error extracting from page {page_num}: {e}")
                    print(f"🔧 Error type: {type(e).__name__}")
                    import traceback
                    print("📋 Tabula error traceback:")
                    traceback.print_exc()
        
        # Clean up OCR file if created
        if ocr_pdf_path and os.path.exists(str(ocr_pdf_path)):
            try:
                os.remove(str(ocr_pdf_path))
                print(f"🧹 Cleaned up OCR file: {ocr_pdf_path}")
            except Exception as cleanup_error:
                print(f"⚠️ Could not clean up OCR file: {cleanup_error}")
        
        if all_tables:
            print(f"\n✅ Successfully extracted {len(all_tables)} tables using ROI selection")
            from pipeline.extract_main import clean_and_filter_tables
            return clean_and_filter_tables(all_tables, "roi")
        else:
            print("\n❌ No tables extracted using ROI selection")
            return []
    
    except Exception as e:
        print(f"❌ Error in ROI-based extraction: {e}")
        import traceback
        traceback.print_exc()
        return []

# Function alias for extract_main.py import
extract_tables_with_roi_selection = extract_tables_with_roi_selection_tabula

if __name__ == "__main__":
    # Add debug output for subprocess execution
    print("🐛 DEBUG: extract_bom_tab.py started as subprocess")
    print(f"🐛 DEBUG: BOM_USE_ROI = {os.environ.get('BOM_USE_ROI', 'NOT SET')}")
    
    # Check if we should run ROI extraction specifically
    if os.environ.get("BOM_USE_ROI", "false").lower() == "true":
        print("🎯 Running ROI-specific extraction...")
        
        # Get configuration from environment
        pdf_path = os.environ.get("BOM_PDF_PATH")
        pages = os.environ.get("BOM_PAGE_RANGE", "all")
        
        if not pdf_path:
            print("❌ No PDF path provided")
            sys.exit(1)
        
        if not os.path.exists(pdf_path):
            print(f"❌ PDF file not found: {pdf_path}")
            sys.exit(1)
        
        try:
            # Run ROI extraction
            print("🔧 Starting ROI extraction...")
            
            # Debug: Show the PDF being processed
            print(f"🐛 DEBUG: extract_bom_tab.py processing PDF: {pdf_path}")
            print(f"🐛 DEBUG: PDF exists: {os.path.exists(pdf_path)}")
            
            # Check if this is an OCR-processed PDF
            is_ocr_pdf = "_ocr" in Path(pdf_path).stem
            print(f"🐛 DEBUG: Is OCR-processed PDF: {is_ocr_pdf}")
            
            if is_ocr_pdf:
                print(f"🔍 DEBUG: OCR-processed PDF detected - tabula ROI extraction will be skipped")
                print(f"🔍 DEBUG: This should return empty tables to trigger Camelot fallback")
            else:
                print(f"🔍 DEBUG: Original PDF - tabula ROI extraction will attempt normally")
            
            tables = extract_tables_with_roi_selection_tabula(pdf_path, pages)
            print(f"🔧 ROI extraction completed, found {len(tables) if tables else 0} tables")
            
            if tables:
                print(f"✅ ROI extraction found {len(tables)} tables")
                
                # Show table selection interface like the normal workflow (always show, even for single table)
                if len(tables) >= 1:
                    if len(tables) > 1:
                        print("📋 Multiple tables found - showing selection interface...")
                    else:
                        print("📋 Single table found - showing selection interface...")
                    from gui.table_selector import show_table_selector
                    selected_tables = show_table_selector(tables)
                    
                    if not selected_tables:
                        print("❌ No tables selected by user")
                        sys.stdout.flush()
                        sys.exit(1)
                else:
                    print("📋 No tables found - cannot proceed")
                    selected_tables = tables
                
                # Apply customer formatting like the main workflow
                print("🔧 Importing functions from main pipeline...")
                from pipeline.extract_main import merge_tables_and_export, save_tables_to_excel, process_and_format_tables
                print("🔧 Functions imported successfully")
                
                # Get company name from environment for customer formatting
                company = os.environ.get("BOM_COMPANY", "")
                print(f"🔧 Company for formatting: '{company}'")
                
                # Process and format tables with customer-specific formatting
                print("🔧 Applying customer formatting...")
                processed_tables = process_and_format_tables(selected_tables, company)
                
                if not processed_tables:
                    print("❌ No tables passed processing/formatting")
                    sys.stdout.flush()
                    sys.exit(1)
                
                print(f"✅ Customer formatting applied - {len(processed_tables)} tables processed")
                
                # Generate output paths
                pdf_dir = Path(pdf_path).parent
                pdf_name = Path(pdf_path).stem
                extracted_path = pdf_dir / f"{pdf_name}_extracted.xlsx"
                merged_path = pdf_dir / f"{pdf_name}_merged.xlsx"
                
                print(f"🔧 Output paths debug:")
                print(f"   PDF path: {pdf_path}")
                print(f"   PDF dir: {pdf_dir}")
                print(f"   PDF name: {pdf_name}")
                print(f"   Extracted: {extracted_path}")
                print(f"   Merged: {merged_path}")
                print(f"   PDF dir exists: {pdf_dir.exists()}")
                print(f"   PDF dir writable: {os.access(pdf_dir, os.W_OK)}")
                
                # Save individual extracted tables first (preserves original table structure)
                print(f"Saving individual extracted tables to: {extracted_path}")
                extracted_success = save_tables_to_excel(selected_tables, str(extracted_path))
                
                if extracted_success:
                    print(f"✅ Individual tables saved: {extracted_path}")
                else:
                    print(f"❌ Failed to save individual tables to: {extracted_path}")
                
                # Save merged tables (combines all tables into one)
                print(f"💾 Saving merged table to: {merged_path}")
                print(f"🔧 Using processed_tables for merged file (count: {len(processed_tables)})")
                
                try:
                    merged_success = merge_tables_and_export(processed_tables, str(merged_path))
                    print(f"🔧 merge_tables_and_export returned: {merged_success}")
                except Exception as merge_error:
                    print(f"❌ merge_tables_and_export failed with exception: {merge_error}")
                    import traceback
                    traceback.print_exc()
                    merged_success = False
                
                print(f"🔧 Final results:")
                print(f"   extracted_success = {extracted_success}")
                print(f"   merged_success = {merged_success}")
                
                if merged_success:
                    print(f"✅ Merged table saved: {merged_path}")
                    print(f"   File exists: {merged_path.exists()}")
                    if merged_path.exists():
                        print(f"   File size: {merged_path.stat().st_size} bytes")
                    
                    if extracted_success:
                        print("✅ ROI extraction completed successfully - both files created")
                        sys.stdout.flush()  # Ensure output is displayed
                        sys.exit(0)
                    else:
                        print("⚠️ ROI extraction partially successful - merged file created but extracted file failed")
                        sys.stdout.flush()  # Ensure output is displayed
                        sys.exit(0)  # Still exit successfully since merged file was created
                else:
                    print("❌ Failed to save merged table")
                    print(f"   merged_success = {merged_success}")
                    print(f"   This is the reason for sys.exit(1)")
                    sys.stdout.flush()  # Ensure output is displayed
                    sys.exit(1)
            else:
                print("❌ No tables found with ROI extraction")
                print("🔄 Tabula ROI extraction complete - returning to orchestrator for Camelot fallback")
                sys.stdout.flush()  # Ensure output is displayed
                sys.exit(0)  # Exit with success code to allow orchestrator to try Camelot
        
        except Exception as e:
            print(f"❌ ROI extraction failed: {e}")
            import traceback
            traceback.print_exc()
            sys.stdout.flush()  # Ensure output is displayed
            sys.exit(1)
    else:
        # Run full extraction workflow
        print("🔄 Running full extraction workflow...")
        from pipeline.extract_main import run_main_extraction_workflow
        run_main_extraction_workflow()
