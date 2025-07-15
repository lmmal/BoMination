import tabula
import pandas as pd
import os
import sys
import re
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from pathlib import Path
from pipeline.validation_utils import validate_extracted_tables, handle_common_errors, generate_output_path
from omni_cust.customer_formatters import apply_customer_formatter, get_available_customers
from gui.table_selector import show_table_selector
from gui.review_window import review_and_edit_dataframe_cli
import logging
from pandastable import Table
import tempfile

# Configure matplotlib to use non-interactive backend to avoid GUI issues
try:
    import matplotlib
    matplotlib.use('Agg')
except ImportError:
    pass  # matplotlib not available

# Import OCR preprocessing module
from pipeline.ocr_preprocessor import (
    preprocess_pdf_with_ocr, 
    cleanup_ocr_temp_files,
    check_ocrmypdf_installation,
    get_ocr_installation_instructions
)

# Configure logging to suppress verbose third-party output
logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')
logging.getLogger('selenium').setLevel(logging.WARNING)
logging.getLogger('urllib3').setLevel(logging.WARNING)
logging.getLogger('requests').setLevel(logging.WARNING)

# Configure Tabula for PyInstaller
if getattr(sys, 'frozen', False):
    # Running as PyInstaller executable
    import tempfile
    # Set java encoding to handle special characters
    os.environ['JAVA_TOOL_OPTIONS'] = '-Dfile.encoding=UTF-8'
    # Set a reasonable temp directory for Tabula
    os.environ['TMPDIR'] = tempfile.gettempdir()

def is_likely_bom_table(df):
    """
    Determine if a dataframe looks like a Bill of Materials table.
    Returns True if it looks like a BOM, False otherwise.
    """
    if df.empty or df.shape[0] < 3 or df.shape[1] < 3:
        return False
    
    # Convert to string for analysis
    df_str = df.astype(str)
    
    # Look for common BOM indicators
    bom_keywords = [
        'ITEM', 'QTY', 'QUANTITY', 'PART', 'NUMBER', 'DESCRIPTION', 'DESC', 
        'MFG', 'MANUFACTURER', 'MPN', 'P/N', 'PART NUMBER', 'ITEM NO',
        'BILL OF MATERIAL', 'BOM', 'COMPONENT', 'REFERENCE', 'REF'
    ]
    
    # Check headers (first few rows)
    header_score = 0
    for i in range(min(3, len(df_str))):
        row_text = ' '.join(df_str.iloc[i].astype(str)).upper()
        for keyword in bom_keywords:
            if keyword in row_text:
                header_score += 1
    
    # Check for structured data patterns
    structure_score = 0
    
    # Look for numeric patterns (item numbers, quantities)
    numeric_cols = 0
    for col in df.columns:
        col_data = df[col].astype(str).str.strip()
        numeric_count = sum(1 for val in col_data if val.isdigit() or val.replace('.', '').isdigit())
        if numeric_count > len(col_data) * 0.3:  # 30% of column is numeric
            numeric_cols += 1
    
    if numeric_cols >= 1:
        structure_score += 2
    
    # Check for consistent row structure (most rows have similar number of filled cells)
    filled_counts = []
    for i in range(len(df)):
        filled_count = sum(1 for val in df.iloc[i].astype(str) if val.strip() != '')
        filled_counts.append(filled_count)
    
    if len(filled_counts) > 0:
        avg_filled = sum(filled_counts) / len(filled_counts)
        consistent_rows = sum(1 for count in filled_counts if abs(count - avg_filled) <= 2)
        if consistent_rows > len(filled_counts) * 0.7:  # 70% of rows are consistent
            structure_score += 1
    
    # Check for strong positive BOM indicators
    all_text = ' '.join(df_str.fillna('').astype(str).values.flatten()).upper()
    
    # Strong positive indicators that this contains BOM data
    strong_bom_indicators = [
        'BILL OF MATERIAL', 'ITEM NO', 'MFG P/N', 'PROTON P/N', 'ALPHA WIRE',
        'HEYCO', 'SIEMENS', 'DELPHI', 'THOMAS BETTS', 'ALTECH', 'MURR'
    ]
    
    strong_positive_score = sum(1 for indicator in strong_bom_indicators if indicator in all_text)
    
    # Reject tables that are clearly not BOMs (but only if they don't have strong BOM indicators)
    reject_keywords = [
        'PRINTED DRAWING', 'REFERENCE ONLY', 'DOCUMENT CONTROL', 'LATEST REVISION',
        'PROPERTY OF', 'DELIVERED ON', 'EXPRESS CONDITION', 'NOT TO BE DISCLOSED',
        'CUT BACK', 'REMOVE', 'SHRINK TUBING', 'DRAWING NUMBER', 'HARNESS'
    ]
    
    reject_score = sum(1 for keyword in reject_keywords if keyword in all_text)
    
    # Decision logic - if we have strong BOM indicators, be more lenient about reject keywords
    if strong_positive_score >= 3:
        # This definitely contains BOM data, even if it has drawing instructions mixed in
        total_score = header_score + structure_score + strong_positive_score - min(reject_score, 5) # Cap reject penalty
        min_threshold = 5  # Lower threshold when we have strong BOM indicators
    else:
        # Standard scoring for tables without strong BOM indicators
        total_score = header_score + structure_score - reject_score
        min_threshold = 2
    
    print(f"    BOM Analysis - Header score: {header_score}, Structure score: {structure_score}, Strong BOM score: {strong_positive_score}, Reject score: {reject_score}, Total: {total_score}, Threshold: {min_threshold}")
    
    # Accept if we meet the threshold
    return total_score >= min_threshold

def detect_pdf_type(pdf_path):
    """
    Detect if a PDF is primarily image-based or text-based.
    Returns 'text' if PDF has searchable text, 'image' if it's primarily image-based.
    """
    try:
        from pipeline.ocr_preprocessor import is_pdf_searchable
        
        print(f"🔍 Analyzing PDF type: {Path(pdf_path).name}")
        
        is_searchable = is_pdf_searchable(pdf_path)
        
        if is_searchable:
            print("� PDF appears to be text-based (has searchable text)")
            return 'text'
        else:
            print("🖼️ PDF appears to be image-based (no searchable text)")
            return 'image'
            
    except Exception as e:
        print(f"⚠️ Could not determine PDF type: {e}")
        print("📄 Assuming text-based PDF")
        return 'text'

def extract_tables_with_tabula_method(pdf_path, pages):
    """
    Extract tables using Tabula-py with multiple fallback methods.
    Works best with text-based PDFs.
    """
    print("📊 Attempting table extraction with Tabula...")
    
    # Set environment variables to handle encoding issues
    original_java_options = os.environ.get('JAVA_TOOL_OPTIONS', '')
    original_lang = os.environ.get('LANG', '')
    original_lc_all = os.environ.get('LC_ALL', '')
    
    try:
        # Set environment to handle subprocess encoding issues
        os.environ['PYTHONIOENCODING'] = 'utf-8:ignore'
        os.environ['LANG'] = 'C.UTF-8'
        os.environ['LC_ALL'] = 'C.UTF-8'
        
        tables = []
        
        # Method 1: Try lattice method with UTF-8 encoding
        try:
            print("  Trying lattice method with UTF-8 encoding...")
            os.environ['JAVA_TOOL_OPTIONS'] = '-Dfile.encoding=UTF-8 -Duser.language=en -Duser.country=US'
            
            tables = tabula.read_pdf(
                pdf_path,
                pages=pages,
                multiple_tables=True,
                lattice=True,
                java_options="-Dfile.encoding=UTF-8 -Duser.language=en -Duser.country=US -Djava.awt.headless=true",
                pandas_options={'header': None}
            )
            print(f"    ✅ Lattice method extracted {len(tables)} tables")
        except Exception as e:
            print(f"    ❌ Lattice method failed: {e}")
            
            # Method 2: Try with different encoding
            try:
                print("  Trying lattice method with ISO-8859-1 encoding...")
                os.environ['PYTHONIOENCODING'] = 'latin1:ignore'
                tables = tabula.read_pdf(
                    pdf_path,
                    pages=pages,
                    multiple_tables=True,
                    lattice=True,
                    java_options="-Dfile.encoding=ISO-8859-1 -Duser.language=en -Duser.country=US -Djava.awt.headless=true",
                    pandas_options={'header': None}
                )
                print(f"    ✅ Lattice method with ISO encoding extracted {len(tables)} tables")
            except Exception as e2:
                print(f"    ❌ Lattice method with ISO encoding failed: {e2}")
                
                # Method 3: Try stream method as final fallback
                try:
                    print("  Trying stream method as final fallback...")
                    os.environ['PYTHONIOENCODING'] = 'utf-8:ignore'
                    tables = tabula.read_pdf(
                        pdf_path,
                        pages=pages,
                        multiple_tables=True,
                        stream=True,
                        guess=True
                    )
                    print(f"    ✅ Stream method extracted {len(tables)} tables")
                except Exception as e3:
                    print(f"    ❌ Stream method failed: {e3}")
                    tables = []
        
        return tables
        
    finally:
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

def extract_tables_with_camelot_method(pdf_path, pages):
    """
    Extract tables using Camelot with both lattice and stream methods.
    Works well with both text-based and image-based PDFs.
    """
    print("🐪 Attempting table extraction with Camelot...")
    
    try:
        import camelot
        
        tables = []
        
        # Method 1: Try lattice method first with better settings
        try:
            print("  Trying Camelot lattice method...")
            tables_camelot = camelot.read_pdf(
                pdf_path, 
                pages=str(pages), 
                flavor='lattice',
                split_text=True,        # Split text in cells
                flag_size=True,         # Use table size for filtering
                strip_text='\n',        # Strip newlines
                line_scale=40,          # Increase line detection sensitivity
                copy_text=['h', 'v'],   # Copy text from horizontal and vertical lines
                shift_text=['l', 't'],  # Shift text left and top
                table_areas=None,       # Auto-detect table areas
                columns=None,           # Auto-detect columns
                process_background=True # Process background for better detection
            )
            
            if tables_camelot and len(tables_camelot) > 0:
                # Filter tables by accuracy and size with structure validation
                good_tables = []
                for i, table in enumerate(tables_camelot):
                    accuracy = table.accuracy
                    rows, cols = table.df.shape
                    print(f"    Camelot table {i+1}: accuracy={accuracy:.1f}%, size={rows}x{cols}")
                    
                    # Check if table has reasonable structure
                    structure_ok = validate_table_structure(table.df)
                    
                    # More balanced filtering - higher accuracy threshold with structure validation
                    if accuracy > 20 and rows >= 5 and cols >= 3 and structure_ok:
                        good_tables.append(table.df)
                        print(f"      ✅ Keeping table {i+1} (good accuracy and structure)")
                    else:
                        print(f"      ❌ Skipping table {i+1} (low accuracy, bad structure, or too small)")
                        print(f"        - Accuracy: {accuracy:.1f}% (need >20%)")
                        print(f"        - Size: {rows}x{cols} (need ≥5x3)")
                        print(f"        - Structure OK: {structure_ok}")
                
                if good_tables:
                    tables = good_tables
                    print(f"    ✅ Camelot lattice extracted {len(tables)} good tables")
                else:
                    print("    ❌ No good quality tables found with lattice method")
            
        except Exception as e:
            print(f"    ❌ Camelot lattice method failed: {e}")
        
        # Method 2: If lattice didn't work, try stream method with better settings
        if not tables:
            try:
                print("  Trying Camelot stream method...")
                tables_camelot = camelot.read_pdf(
                    pdf_path, 
                    pages=str(pages), 
                    flavor='stream',
                    split_text=True,        # Split text in cells
                    flag_size=True,         # Use table size for filtering
                    strip_text='\n',        # Strip newlines
                    edge_tol=500,           # Tolerance for edge detection
                    row_tol=2,             # Row tolerance for better separation
                    column_tol=0           # Column tolerance for better separation
                )
                
                if tables_camelot and len(tables_camelot) > 0:
                    # Filter tables by accuracy and size with structure validation
                    good_tables = []
                    for i, table in enumerate(tables_camelot):
                        accuracy = table.accuracy
                        rows, cols = table.df.shape
                        print(f"    Camelot table {i+1}: accuracy={accuracy:.1f}%, size={rows}x{cols}")
                        
                        # Check if table has reasonable structure
                        structure_ok = validate_table_structure(table.df)
                        
                        # Lower threshold for stream method but still validate structure
                        if accuracy > 15 and rows >= 3 and cols >= 3 and structure_ok:
                            good_tables.append(table.df)
                            print(f"      ✅ Keeping table {i+1} (acceptable accuracy and structure)")
                        else:
                            print(f"      ❌ Skipping table {i+1} (low accuracy, bad structure, or too small)")
                            print(f"        - Accuracy: {accuracy:.1f}% (need >15%)")
                            print(f"        - Size: {rows}x{cols} (need ≥3x3)")
                            print(f"        - Structure OK: {structure_ok}")
                    
                    if good_tables:
                        tables = good_tables
                        print(f"    ✅ Camelot stream extracted {len(tables)} good tables")
                    else:
                        print("    ❌ No good quality tables found with stream method")
                        
            except Exception as e:
                print(f"    ❌ Camelot stream method failed: {e}")
        
        # Method 3: If both methods failed, try with very lenient settings as last resort
        if not tables:
            print("  Trying Camelot with very lenient settings as last resort...")
            try:
                tables_camelot = camelot.read_pdf(
                    pdf_path, 
                    pages=str(pages), 
                    flavor='lattice',
                    split_text=True,
                    flag_size=False,        # Don't use size filtering
                    strip_text='\n',
                    line_scale=15,          # Very sensitive line detection
                    copy_text=['h', 'v'],
                    shift_text=['l', 't'],
                    process_background=False
                )
                
                if tables_camelot and len(tables_camelot) > 0:
                    # Accept any table that has reasonable dimensions
                    for i, table in enumerate(tables_camelot):
                        accuracy = table.accuracy
                        rows, cols = table.df.shape
                        print(f"    Camelot table {i+1}: accuracy={accuracy:.1f}%, size={rows}x{cols}")
                        
                        # Very lenient - just check minimum size
                        if rows >= 3 and cols >= 3:
                            # Try to repair the table structure
                            repaired_table = repair_table_structure(table.df)
                            if repaired_table is not None:
                                tables.append(repaired_table)
                                print(f"      ✅ Keeping table {i+1} (repaired structure)")
                            else:
                                print(f"      ❌ Could not repair table {i+1} structure")
                        else:
                            print(f"      ❌ Skipping table {i+1} (too small)")
                    
                    if tables:
                        print(f"    ✅ Camelot extracted {len(tables)} tables with lenient settings")
                    
            except Exception as e:
                print(f"    ❌ Camelot lenient method failed: {e}")
        
        if not tables:
            print("❌ Camelot found no suitable tables")
            
        return tables
        
    except ImportError:
        print("❌ Camelot not available - install with: pip install camelot-py[cv]")
        return []
    except Exception as e:
        print(f"❌ Camelot extraction failed: {e}")
        return []

def validate_table_structure(df):
    """
    Validate if a table has reasonable structure for BOM data.
    Returns True if the table structure looks reasonable.
    """
    try:
        if df is None or df.empty:
            return False
        
        rows, cols = df.shape
        
        # Check if table has reasonable dimensions
        if rows < 3 or cols < 3:
            return False
        
        # Check if cells are not mostly empty
        non_empty_cells = df.astype(str).apply(lambda x: x.str.strip() != '').sum().sum()
        total_cells = rows * cols
        fill_ratio = non_empty_cells / total_cells
        
        if fill_ratio < 0.1:  # Less than 10% of cells have content
            return False
        
        # Check if there's not too much text crammed into single cells
        # (indicates poor column separation)
        avg_cell_length = df.astype(str).apply(lambda x: x.str.len()).mean().mean()
        if avg_cell_length > 200:  # Very long average cell content
            return False
        
        # Check if first few rows don't have extremely long merged content
        for i in range(min(3, rows)):
            for j in range(cols):
                cell_content = str(df.iloc[i, j]).strip()
                if len(cell_content) > 500:  # Single cell with >500 characters
                    return False
        
        return True
        
    except Exception as e:
        print(f"⚠️ Table structure validation failed: {e}")
        return False

def repair_table_structure(df):
    """
    Attempt to repair poorly structured tables by splitting merged content.
    Returns repaired DataFrame or None if repair fails.
    """
    try:
        if df is None or df.empty:
            return None
        
        # Create a copy to work with
        repaired_df = df.copy()
        
        # Look for cells with multiple part numbers/descriptions merged together
        # This is a basic repair - could be enhanced further
        for i in range(len(repaired_df)):
            for j in range(len(repaired_df.columns)):
                cell_content = str(repaired_df.iloc[i, j]).strip()
                
                # If cell contains multiple part numbers (pattern: numbers followed by text)
                if len(cell_content) > 100:
                    # Try to split on common patterns
                    import re
                    # Pattern for part numbers like "001 SIEMENS" or "002 RITTAL"
                    parts = re.split(r'(\d{3,}\s+[A-Z][A-Z\s]+)', cell_content)
                    if len(parts) > 3:  # Found multiple parts
                        # Keep only the first part for this cell
                        repaired_df.iloc[i, j] = parts[1] if len(parts) > 1 else cell_content[:100]
        
        # Validate the repaired table
        if validate_table_structure(repaired_df):
            return repaired_df
        else:
            return None
            
    except Exception as e:
        print(f"⚠️ Table repair failed: {e}")
        return None

def process_pdf_with_ocr(pdf_path, force_ocr=False):
    """
    Process a PDF with OCR to make it searchable.
    Returns the path to the OCR'd PDF or None if OCR failed.
    """
    print(f"🔄 Processing PDF with OCR (force_ocr={force_ocr})...")
    
    try:
        ocr_available, ocr_version, ocr_error = check_ocrmypdf_installation()
        if not ocr_available:
            print(f"❌ OCR not available: {ocr_error}")
            return None
        
        print(f"✅ OCR available: {ocr_version}")
        
        # Create output path in same directory as input PDF
        pdf_path_obj = Path(pdf_path)
        output_path = pdf_path_obj.parent / f"{pdf_path_obj.stem}_ocr{pdf_path_obj.suffix}"
        
        # Process with OCR
        ocr_success, ocr_pdf_path, ocr_error = preprocess_pdf_with_ocr(
            pdf_path, 
            output_path=str(output_path), 
            force_ocr=force_ocr
        )
        
        if ocr_success:
            print(f"✅ OCR processing successful!")
            print(f"📄 OCR'd PDF: {ocr_pdf_path}")
            
            # Check if the OCR'd PDF is actually searchable now
            if not force_ocr:
                try:
                    from pipeline.ocr_preprocessor import is_pdf_searchable
                    is_searchable = is_pdf_searchable(ocr_pdf_path)
                    
                    if not is_searchable:
                        print("⚠️ OCR'd PDF still not searchable, retrying with --force-ocr...")
                        # Clean up the non-searchable result
                        cleanup_ocr_temp_files(ocr_pdf_path)
                        # Retry with force OCR
                        return process_pdf_with_ocr(pdf_path, force_ocr=True)
                except Exception as e:
                    print(f"⚠️ Could not verify OCR'd PDF searchability: {e}")
            
            return ocr_pdf_path
        else:
            print(f"❌ OCR processing failed: {ocr_error}")
            return None
            
    except Exception as e:
        print(f"❌ OCR processing error: {e}")
        return None

def clean_and_filter_tables(tables, method_name):
    """
    Clean and filter extracted tables to keep only those that look like BOMs.
    """
    print(f"🧹 Cleaning and filtering {len(tables)} tables from {method_name}...")
    
    cleaned_tables = []
    for i, table in enumerate(tables):
        if table is not None and not table.empty:
            # Convert all data to string and handle NaN values properly
            table = table.fillna('')  # Fill NaN with empty strings first
            table = table.astype(str)  # Convert to string
            
            # Clean up common extraction artifacts
            table = table.replace('nan', '')
            table = table.replace('None', '')
            
            # Remove completely empty rows and columns
            table = table.loc[:, (table != '').any(axis=0)]  # Remove empty columns
            table = table.loc[(table != '').any(axis=1)]     # Remove empty rows
            
            # Reset index
            table = table.reset_index(drop=True)
            
            # Clean individual cells of problematic characters
            for col in table.columns:
                # More aggressive cleaning for better readability
                table[col] = table[col].str.replace(r'[^\w\s\-\.\,\/\(\)\:]', ' ', regex=True)  # Keep only alphanumeric, spaces, and common punctuation
                table[col] = table[col].str.replace(r'\s+', ' ', regex=True)  # Normalize whitespace
                table[col] = table[col].str.strip()  # Remove leading/trailing whitespace
            
            # Filter out tables that don't look like BOMs
            if not table.empty:
                # Check if this looks like a BOM table
                is_bom_table = is_likely_bom_table(table)
                if is_bom_table:
                    cleaned_tables.append(table)
                    print(f"  Table {i+1}: {table.shape[0]} rows, {table.shape[1]} columns - ✅ Looks like BOM")
                    # Show a sample of the cleaned data for debugging
                    print(f"  Sample cleaned data from Table {i+1}:")
                    print(table.head(2).to_string())
                    print("  ---")
                else:
                    print(f"  Table {i+1}: {table.shape[0]} rows, {table.shape[1]} columns - ❌ Doesn't look like BOM, skipping")
                    # Show sample of rejected table for debugging
                    print(f"  Sample rejected data from Table {i+1}:")
                    print(table.head(2).to_string())
                    print("  ---")
    
    return cleaned_tables

def extract_tables_from_pdf(pdf_path, pages):
    """
    Main function to extract tables from PDF using the best strategy based on PDF type.
    """
    ocr_pdf_path = None  # Track OCR'd PDF for cleanup
    
    try:
        print(f"\n🔍 Extracting tables from pages: {pages}")
        print(f"📄 PDF path: {pdf_path}")
        
        # Handle path encoding issues
        try:
            # Ensure the PDF path is properly encoded
            if isinstance(pdf_path, str):
                pdf_path = pdf_path.encode('utf-8', errors='ignore').decode('utf-8')
            print(f"📄 Normalized PDF path: {pdf_path}")
        except Exception as path_error:
            print(f"⚠️ Warning: Could not normalize PDF path: {path_error}")
        
        # Check if running in PyInstaller environment
        if getattr(sys, 'frozen', False):
            print("🔧 Running in PyInstaller environment")
            print(f"🔧 JAVA_TOOL_OPTIONS: {os.environ.get('JAVA_TOOL_OPTIONS', 'Not set')}")
        
        # Step 1: Detect PDF type
        pdf_type = detect_pdf_type(pdf_path)
        
        tables = []
        
        if pdf_type == 'text':
            print("\n📊 PDF is text-based - trying text extraction methods...")
            
            # Try Tabula first (best for text-based PDFs)
            tables = extract_tables_with_tabula_method(pdf_path, pages)
            
            if not tables:
                print("📊 Tabula failed - trying Camelot as fallback...")
                tables = extract_tables_with_camelot_method(pdf_path, pages)
                
        else:  # pdf_type == 'image'
            print("\n🖼️ PDF is image-based - processing with OCR first...")
            
            # For image-based PDFs, start with force_ocr=True to handle vector content
            # Vector PDFs need force_ocr to be processed properly
            ocr_pdf_path = process_pdf_with_ocr(pdf_path, force_ocr=True)
            
            if ocr_pdf_path:
                print("📊 OCR successful - trying table extraction on OCR'd PDF...")
                working_pdf_path = ocr_pdf_path
                
                # Try Camelot first (often works better with OCR'd PDFs)
                tables = extract_tables_with_camelot_method(working_pdf_path, pages)
                
                if not tables:
                    print("📊 Camelot failed - trying Tabula as fallback...")
                    tables = extract_tables_with_tabula_method(working_pdf_path, pages)
            else:
                print("❌ OCR failed - falling back to original PDF with text extraction methods...")
                # Fallback to original PDF with text methods
                tables = extract_tables_with_tabula_method(pdf_path, pages)
                
                if not tables:
                    tables = extract_tables_with_camelot_method(pdf_path, pages)
        
        # Step 2: Clean and filter tables
        if tables:
            print(f"\n🧹 Processing {len(tables)} extracted tables...")
            
            # Step 2a: Check for and split dual-column BOM tables (only for certain patterns)
            processed_tables = []
            for i, table in enumerate(tables):
                print(f"  Processing table {i+1}...")
                # Only apply dual-column splitting if it looks like a dual-column BOM
                if should_split_dual_column(table):
                    split_table = split_dual_column_bom(table)
                    processed_tables.append(split_table)
                else:
                    processed_tables.append(table)
            
            # Step 2b: Clean and filter the processed tables
            method_name = "Tabula" if pdf_type == 'text' else "Camelot"
            cleaned_tables = clean_and_filter_tables(processed_tables, method_name)
            
            if cleaned_tables:
                print(f"🎉 Successfully extracted {len(cleaned_tables)} BOM tables")
                
                # Step 3: Validate extracted tables
                tables_valid, warning_message = validate_extracted_tables(cleaned_tables)
                
                if not tables_valid:
                    print(warning_message)
                    # Show warning in GUI if available
                    try:
                        messagebox.showwarning("Table Extraction Warning", warning_message)
                    except:
                        pass  # GUI not available, warning already printed
                    return []
                elif warning_message:
                    print(warning_message)
                    try:
                        messagebox.showinfo("Table Extraction Info", warning_message)
                    except:
                        pass  # GUI not available, info already printed
                
                print(f"✅ Final result: {len(cleaned_tables)} valid BOM tables extracted")
                return cleaned_tables
            else:
                print("❌ No valid BOM tables found after cleaning and filtering")
                return []
        else:
            print("❌ No tables extracted from the PDF")
            return []
        
    except Exception as e:
        error_message = f"Error during table extraction: {e}"
        friendly_error = handle_common_errors(str(e))
        print(friendly_error)
        
        # Show error in GUI if available
        try:
            messagebox.showerror("Table Extraction Error", friendly_error)
        except:
            pass  # GUI not available, error already printed
        
        return []
    
    finally:
        # Clean up OCR temporary files only if they are in temp directories
        if ocr_pdf_path:
            print("🧹 Cleaning up OCR temporary files...")
            # Only clean up if it's in a temp directory (not our permanent OCR files)
            if "bomination_ocr_" in str(ocr_pdf_path):
                cleanup_ocr_temp_files(ocr_pdf_path)
            else:
                print("📁 OCR'd PDF saved permanently, not cleaning up")

# Keep the old function name for backward compatibility
def extract_tables_with_tabula(pdf_path, pages):
    """
    Legacy function name - redirects to the new main function.
    """
    return extract_tables_from_pdf(pdf_path, pages)

# Customer-specific formatting has been moved to customer_formatters.py
# Use apply_customer_formatter(df, customer_name) instead of calling specific functions directly

def format_table_as_text(table):
    """Format a pandas DataFrame as readable text with proper spacing and wrapping."""
    if table.empty:
        return "Empty table"
    
    # Get column headers
    headers = [f"Col_{j}" for j in range(len(table.columns))]
    
    # Calculate column widths dynamically based on ALL rows
    col_widths = []
    for j in range(len(table.columns)):
        max_width = len(headers[j])
        for idx in range(len(table)):  # Check all rows, not just first few
            cell_value = str(table.iloc[idx, j])
            if cell_value and cell_value.strip():
                # Limit individual cell display to reasonable length
                cell_value = ' '.join(cell_value.split())
                max_width = max(max_width, min(len(cell_value), 40))
        col_widths.append(max(max_width, 10))  # Minimum width of 10
    
    # Format the table text
    text_lines = []
    
    # Header row
    header_line = " | ".join(header.ljust(width) for header, width in zip(headers, col_widths))
    text_lines.append(header_line)
    text_lines.append("-" * len(header_line))
    
    # Display ALL rows
    for idx in range(len(table)):
        row_parts = []
        for col_idx, width in enumerate(col_widths):
            cell_value = str(table.iloc[idx, col_idx])
            if cell_value and cell_value.strip():
                cell_value = ' '.join(cell_value.split())
                # Truncate if too long for the column
                if len(cell_value) > width:
                    cell_value = cell_value[:width-3] + "..."
            else:
                cell_value = ""
            row_parts.append(cell_value.ljust(width))
        
        text_lines.append(" | ".join(row_parts))
    
    return "\n".join(text_lines)

# Table selector functionality has been moved to table_selector.py
# Use show_table_selector(tables) from the imported module

# Review functionality has been moved to review_window.py
# Use review_and_edit_dataframe_cli(df) from the imported module

def save_tables_to_excel(tables, output_path):
    try:
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            for i, table in enumerate(tables):
                sheet_name = f"Table_{i+1}"
                table.to_excel(writer, index=False, sheet_name=sheet_name)
        print(f"\nTables saved to: {output_path}")
    except Exception as e:
        print(f"Failed to save Excel file: {e}")

def merge_tables_and_export(tables, output_path, sheet_name="Combined_BoM", company=""):
    print(f"\n📊 MERGE DEBUG: ===== MERGE_TABLES_AND_EXPORT CALLED =====")
    print(f"📊 MERGE DEBUG: Received {len(tables)} tables to merge")
    for i, table in enumerate(tables):
        print(f"📊 MERGE DEBUG: Table {i+1} for merge: shape={table.shape}, sample={table.iloc[0].to_dict() if len(table) > 0 else 'EMPTY'}")
    print(f"📊 MERGE DEBUG: Output path: {output_path}")
    print(f"📊 MERGE DEBUG: ===== END MERGE INPUT DEBUG =====")
    
    try:
        merged_df = pd.concat(tables, ignore_index=True)
        
        print(f"\n📊 MERGE DEBUG: ===== BEFORE COMPANY PROCESSING =====")
        print(f"📊 MERGE DEBUG: Merged table shape: {merged_df.shape}")
        print(f"📊 MERGE DEBUG: Current columns: {merged_df.columns.tolist()}")
        print(f"📊 MERGE DEBUG: First 3 rows:")
        print(merged_df.head(3).to_string())
        print(f"📊 MERGE DEBUG: ===== END BEFORE PROCESSING =====")

        company = company.lower()
        if company == "farrell":
            print("📊 MERGE DEBUG: Applying Farrell-specific table formatting...")
            merged_df = apply_customer_formatter(merged_df, 'farrell')
        elif company == "nel":
            print("📊 MERGE DEBUG: Applying NEL-specific table formatting...")
            merged_df = apply_customer_formatter(merged_df, 'nel')
        elif company == "primetals":
            print("📊 MERGE DEBUG: Applying Primetals-specific table formatting...")
            merged_df = apply_customer_formatter(merged_df, 'primetals')
        else:
            print(f"📊 MERGE DEBUG: No company-specific formatting applied (company='{company}')")
            
            # Try to auto-detect company from the data
            merged_text = ' '.join(merged_df.fillna('').astype(str).values.flatten()).upper()
            if 'NEL HYDROGEN' in merged_text or 'PROTON ENERGY' in merged_text:
                print("📊 MERGE DEBUG: Auto-detected NEL company from data, applying NEL formatting...")
                merged_df = apply_customer_formatter(merged_df, 'nel')
            elif 'FARRELL' in merged_text:
                print("📊 MERGE DEBUG: Auto-detected Farrell company from data, applying Farrell formatting...")
                merged_df = apply_customer_formatter(merged_df, 'farrell')
            elif 'PRIMETALS' in merged_text:
                print("📊 MERGE DEBUG: Auto-detected Primetals company from data, applying Primetals formatting...")
                merged_df = apply_customer_formatter(merged_df, 'primetals')
            else:
                # Apply generic formatting as fallback
                print("📊 MERGE DEBUG: No customer detected, applying generic formatting...")
                merged_df = apply_customer_formatter(merged_df, 'generic')

        print(f"\n📊 MERGE DEBUG: ===== AFTER COMPANY PROCESSING =====")
        print(f"📊 MERGE DEBUG: Table shape: {merged_df.shape}")
        print(f"📊 MERGE DEBUG: Final columns: {merged_df.columns.tolist()}")
        print(f"📊 MERGE DEBUG: First 3 rows after company processing:")
        print(merged_df.head(3).to_string())
        print(f"📊 MERGE DEBUG: ===== END AFTER PROCESSING =====\n")

        # Remove completely empty rows
        merged_df = merged_df.dropna(how='all')
        
        # Remove any remaining duplicate header rows
        header_row = merged_df.columns.tolist()
        merged_df = merged_df[~merged_df.apply(lambda row: row.tolist() == header_row, axis=1)]
        
        # Remove duplicate rows and reset index
        merged_df = merged_df.drop_duplicates().reset_index(drop=True)
        
        # CRITICAL: Apply final "N/A" filling after all other processing
        # This ensures empty cells are filled with "N/A" for OEMSecrets compatibility
        print("📊 MERGE DEBUG: Applying final 'N/A' filling for OEMSecrets compatibility...")
        merged_df = merged_df.fillna('N/A')
        merged_df = merged_df.replace(['', 'nan', 'None', 'NaN'], 'N/A')
        
        print(f"📊 MERGE DEBUG: Final cleaned table shape: {merged_df.shape}")
        print(f"📊 MERGE DEBUG: Sample of final data with N/A filling:")
        print(merged_df.head(3).to_string())

        # Check if we should show the review window
        # If called from GUI thread, we'll skip the review here and handle it in the GUI
        skip_review = os.environ.get("BOM_SKIP_REVIEW", "false").lower() == "true"
        
        if not skip_review:
            print("Opening manual review window...")
            merged_df = review_and_edit_dataframe_cli(merged_df)
        else:
            print("Skipping manual review (will be handled by GUI)...")

        merged_df.to_excel(output_path, index=False, sheet_name=sheet_name)
        print(f"✅ Final cleaned and reviewed table saved to: {output_path}")
        
        # Return the dataframe for GUI processing
        return merged_df

    except Exception as e:
        print(f"❌ Failed to merge and export tables: {e}")
        import traceback
        traceback.print_exc()

def main():
    print("\n=== BoM Table Extractor with GUI Preview ===\n")
    pdf_path = os.environ.get("BOM_PDF_PATH")
    pages = os.environ.get("BOM_PAGE_RANGE")
    company = os.environ.get("BOM_COMPANY")
    output_directory = os.environ.get("BOM_OUTPUT_DIRECTORY")

    if not pdf_path or not pages:
        error_msg = "Missing required input parameters:\n"
        if not pdf_path:
            error_msg += "- BOM_PDF_PATH environment variable not set\n"
        if not pages:
            error_msg += "- BOM_PAGE_RANGE environment variable not set\n"
        print(error_msg)
        return

    print(f"Processing PDF: {pdf_path}")
    print(f"Page range: {pages}")
    print(f"Company: {company or 'None specified'}")
    print(f"Output directory: {output_directory or 'Same as input PDF'}")

    tables = extract_tables_from_pdf(pdf_path, pages)

    if tables:
        print(f"\nSUCCESS: Successfully extracted {len(tables)} table(s)")
        selected_tables = show_table_selector(tables)

        if not selected_tables:
            print("No tables selected. Exiting.")
            return

        # Generate output paths - save to PDF directory for debugging
        # Use PDF directory instead of output_directory parameter
        pdf_dir = Path(pdf_path).parent
        pdf_name = Path(pdf_path).stem
        
        extracted_path = pdf_dir / f"{pdf_name}_extracted.xlsx"
        merged_path = pdf_dir / f"{pdf_name}_merged.xlsx"
        
        print(f"📁 Saving extracted tables to: {extracted_path}")
        print(f"📁 Saving merged table to: {merged_path}")

        save_tables_to_excel(selected_tables, extracted_path)
        merge_tables_and_export(selected_tables, merged_path, company=company)
        
        print(f"\nSUCCESS: Table extraction completed successfully!")
        print(f"OUTPUT: Extracted tables saved to: {extracted_path}")
        print(f"OUTPUT: Merged table saved to: {merged_path}")
    else:
        error_msg = (
            "❌ No tables extracted from the PDF.\n\n"
            "2. Check if the PDF has text-based tables (not images)\n"
            "3. Try a different page range or extraction method\n"
            "4. Consider using OCR for image-based tables\n\n"
            "The pipeline cannot continue without extracted tables."
        )
        print(error_msg)
        
        # Show error in GUI if available
        try:
            import matplotlib
            matplotlib.use('Agg')  # Use non-interactive backend to avoid GUI issues
            messagebox.showerror("No Tables Extracted", error_msg)
        except:
            pass  # GUI not available, error already printed
        
        # Exit with error code to indicate failure
        sys.exit(1)

if __name__ == "__main__":
    main()

def should_split_dual_column(df):
    """
    Determine if a table should be split using dual-column logic.
    
    Args:
        df: DataFrame to check
        
    Returns:
        bool: True if this looks like a dual-column BOM table
    """
    # Check if this looks like a dual-column BOM
    if df.shape[1] < 8:
        return False
    
    # Look for header patterns to identify column groups
    header_row = df.iloc[0] if len(df) > 0 else pd.Series()
    header_str = ' '.join(str(cell) for cell in header_row)
    
    # Find repeated patterns indicating dual columns
    if 'ITEM' in header_str and header_str.count('ITEM') >= 2:
        return True
    
    return False

def split_dual_column_bom(df):
    """
    Split a dual-column BOM table into individual parts.
    
    Args:
        df: DataFrame with side-by-side BOM columns
        
    Returns:
        DataFrame with individual parts in rows
    """
    
    # Check if this looks like a dual-column BOM
    if df.shape[1] < 8:
        return df
    
    # Look for header patterns to identify column groups
    header_row = df.iloc[0] if len(df) > 0 else pd.Series()
    header_str = ' '.join(str(cell) for cell in header_row)
    
    print(f"    Table shape: {df.shape}")
    print(f"    Header row: {header_str}")
    
    # Find repeated patterns indicating dual columns
    if 'ITEM' in header_str and header_str.count('ITEM') >= 2:
        print("🔍 Detected dual-column BOM table - splitting into individual parts")
        
        # Find the column indices for each side
        left_cols = []
        right_cols = []
        
        # For dual-column BOM, we expect: ITEM, MFG, MFGPART, DESCRIPTION, QTY on each side
        header_list = [str(cell).strip().upper() for cell in header_row]
        
        # Find all ITEM columns (there should be 2)
        item_positions = [i for i, h in enumerate(header_list) if 'ITEM' in h]
        print(f"    ITEM positions found: {item_positions}")
        
        if len(item_positions) >= 2:
            # Use the ITEM positions to determine left and right column groups
            left_start = item_positions[0]
            right_start = item_positions[1]
            
            # Standard BOM columns: ITEM, MFG, MFGPART, DESCRIPTION, QTY
            left_cols = list(range(left_start, min(left_start + 5, len(header_list))))
            right_cols = list(range(right_start, min(right_start + 5, len(header_list))))
            
            # Remove any columns that don't exist
            left_cols = [i for i in left_cols if i < len(header_list)]
            right_cols = [i for i in right_cols if i < len(header_list)]
            
            print(f"    Left columns: {left_cols}")
            print(f"    Right columns: {right_cols}")
            
            if len(left_cols) >= 4 and len(right_cols) >= 4:
                # Extract data from both sides
                left_data = df.iloc[:, left_cols].copy()
                right_data = df.iloc[:, right_cols].copy()
                
                # Standardize column names
                standard_cols = ['ITEM', 'MFG', 'MFGPART', 'DESCRIPTION', 'QTY']
                left_data.columns = standard_cols[:len(left_data.columns)]
                right_data.columns = standard_cols[:len(right_data.columns)]
                
                # Remove header rows and empty rows
                left_data = left_data[1:].reset_index(drop=True)  # Skip header
                right_data = right_data[1:].reset_index(drop=True)  # Skip header
                
                # Filter out empty rows
                left_data = left_data[left_data['ITEM'].notna() & (left_data['ITEM'].astype(str).str.strip() != '')].copy()
                right_data = right_data[right_data['ITEM'].notna() & (right_data['ITEM'].astype(str).str.strip() != '')].copy()
                
                print(f"    Left side: {len(left_data)} parts")
                print(f"    Right side: {len(right_data)} parts")
                
                # Combine both sides
                combined_data = pd.concat([left_data, right_data], ignore_index=True)
                
                # Clean up the data
                combined_data = combined_data.dropna(subset=['ITEM'])
                combined_data = combined_data[combined_data['ITEM'].astype(str).str.strip() != '']
                
                print(f"    Combined: {len(combined_data)} parts")
                
                return combined_data
    
    return df