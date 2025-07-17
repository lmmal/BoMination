"""
Main extraction orchestrator that handles the workflow for table extraction from PDFs.
This module coordinates between different extraction methods and OCR preprocessing.
"""

import os
import sys
import logging
import json
from pathlib import Path
import traceback

# Import extraction modules
from pipeline.extract_bom_tab import extract_tables_with_tabula_method_impl, extract_tables_with_roi_selection
from pipeline.extract_bom_cam import extract_tables_with_camelot_method
from pipeline.ocr_preprocessor import process_pdf_with_ocr, cleanup_ocr_temp_files
from pipeline.validation_utils import validate_extracted_tables, handle_common_errors, generate_output_path

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')
logger = logging.getLogger(__name__)

def is_table_good_quality(df):
    """
    Check if a DataFrame represents a good quality table.
    
    This function evaluates table quality based on:
    1. Content density (ratio of non-empty cells)
    2. Text readability (avoiding common OCR artifacts)
    3. Structure consistency (column alignment)
    
    Args:
        df (pandas.DataFrame): Table to evaluate
        
    Returns:
        bool: True if table is good quality, False otherwise
    """
    if df is None or df.empty:
        return False
    
    try:
        # Calculate content density
        total_cells = df.shape[0] * df.shape[1]
        non_empty_cells = df.notna().sum().sum()
        content_density = non_empty_cells / total_cells if total_cells > 0 else 0
        
        # Check for minimum content density (at least 20% filled)
        if content_density < 0.2:
            print(f"    ❌ Quality Check: Low content density ({content_density:.2f})")
            return False
        
        # Check for text readability - look for common OCR artifacts
        poor_quality_indicators = [
            "ALL WORKMANSHIP SHAL",  # Common OCR artifact from your output
            "1.ALL WORKMANSHIP SHAL",
            "WORKMANSHIP SHAL",
            "NOTES:",
            "NOTE:",
            "GENERAL NOTES",
            "ALL WORKMANSHIP SHALL",
            "SHALL BE",
            "SPECIFICATION",
            "SPECIFICATIONS"
        ]
        
        # Convert all text to string and check for artifacts
        all_text = ""
        for col in df.columns:
            col_text = df[col].fillna("").astype(str).str.upper()
            all_text += " ".join(col_text.tolist()) + " "
        
        # Check if the table is mostly composed of poor quality text
        artifact_count = sum(1 for indicator in poor_quality_indicators if indicator in all_text)
        
        if artifact_count > 2:  # More than 2 artifacts suggests poor quality
            print(f"    ❌ Quality Check: Too many OCR artifacts ({artifact_count})")
            return False
        
        # Check for reasonable column structure
        if df.shape[1] < 2:  # Need at least 2 columns for a useful table
            print(f"    ❌ Quality Check: Too few columns ({df.shape[1]})")
            return False
        
        # Check for reasonable row count
        if df.shape[0] < 3:  # Need at least 3 rows (header + 2 data rows)
            print(f"    ❌ Quality Check: Too few rows ({df.shape[0]})")
            return False
        
        # Check for diverse content (not all the same text)
        unique_values = set()
        for col in df.columns:
            col_values = df[col].fillna("").astype(str).str.strip()
            unique_values.update(col_values.tolist())
        
        if len(unique_values) < 5:  # Need at least 5 unique values for diversity
            print(f"    ❌ Quality Check: Low content diversity ({len(unique_values)} unique values)")
            return False
        
        print(f"    ✅ Quality Check: Good quality table (density: {content_density:.2f}, artifacts: {artifact_count})")
        return True
        
    except Exception as e:
        print(f"    ❌ Quality Check: Error evaluating table quality: {e}")
        return False


def visualize_camelot_roi_on_pdf(pdf_path, page_num, camelot_area, original_tabula_area):
    """
    Create a visual debug image showing the ROI area overlaid on the PDF page.
    
    Args:
        pdf_path (str): Path to the PDF file
        page_num (int): Page number (1-indexed)
        camelot_area (list): [x1, y1, x2, y2] in Camelot format (PDF coordinate space)
        original_tabula_area (list): [top, left, bottom, right] in Tabula format
        
    Returns:
        str: Path to the debug image file, or None if failed
    """
    try:
        # Import required libraries
        import fitz  # PyMuPDF
        from PIL import Image, ImageDraw, ImageFont
        import io
        
        # Open PDF
        pdf_doc = fitz.open(pdf_path)
        page = pdf_doc[page_num - 1]  # Convert to 0-indexed
        
        # Get page dimensions
        page_rect = page.rect
        page_width = page_rect.width
        page_height = page_rect.height
        
        print(f"  📏 PDF Page dimensions: {page_width:.1f} x {page_height:.1f} points")
        
        # Convert PDF page to image
        zoom = 2.0  # Higher resolution for better visibility
        mat = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat)
        img_data = pix.tobytes("ppm")
        
        # Create PIL image
        image = Image.open(io.BytesIO(img_data))
        draw = ImageDraw.Draw(image)
        
        # Scale coordinates to image resolution
        x1, y1, x2, y2 = camelot_area
        
        # In PDF coordinate space: (0,0) is bottom-left
        # Convert to image coordinate space: (0,0) is top-left
        img_x1 = x1 * zoom
        img_y1 = (page_height - y2) * zoom  # Flip Y coordinate
        img_x2 = x2 * zoom
        img_y2 = (page_height - y1) * zoom  # Flip Y coordinate
        
        print(f"  📐 Camelot area in PDF coords: x1={x1:.1f}, y1={y1:.1f}, x2={x2:.1f}, y2={y2:.1f}")
        print(f"  📐 Camelot area in image coords: x1={img_x1:.1f}, y1={img_y1:.1f}, x2={img_x2:.1f}, y2={img_y2:.1f}")
        
        # Draw the ROI rectangle
        rectangle_coords = [img_x1, img_y1, img_x2, img_y2]
        draw.rectangle(rectangle_coords, outline="red", width=4)
        
        # Add labels
        try:
            # Try to load a font, fallback to default if not available
            font = ImageFont.truetype("arial.ttf", 24)
        except:
            font = ImageFont.load_default()
        
        # Add coordinate labels
        label_text = f"Camelot ROI: ({x1:.0f},{y1:.0f}) to ({x2:.0f},{y2:.0f})"
        draw.text((img_x1, img_y1 - 30), label_text, fill="red", font=font)
        
        # Add original tabula coordinates for comparison
        tabula_text = f"Original Tabula: [top={original_tabula_area[0]:.0f}, left={original_tabula_area[1]:.0f}, bottom={original_tabula_area[2]:.0f}, right={original_tabula_area[3]:.0f}]"
        draw.text((img_x1, img_y2 + 10), tabula_text, fill="blue", font=font)
        
        # Add coordinate system info
        coord_info = "PDF coords: (0,0) = bottom-left, Image coords: (0,0) = top-left"
        draw.text((10, 10), coord_info, fill="green", font=font)
        
        # Save debug image
        pdf_stem = Path(pdf_path).stem
        debug_filename = f"{pdf_stem}_page{page_num}_camelot_roi_debug.png"
        debug_path = Path(pdf_path).parent / debug_filename
        
        image.save(debug_path)
        pdf_doc.close()
        
        return str(debug_path)
        
    except ImportError as e:
        print(f"  ⚠️ ROI visualization requires PyMuPDF and PIL: {e}")
        return None
    except Exception as e:
        print(f"  ⚠️ ROI visualization error: {e}")
        return None


def clean_and_filter_tables(tables, method_name):
    """
    Clean extracted tables and filter for reasonable candidates.
    
    Args:
        tables (list): List of DataFrames to clean
        method_name (str): Name of the extraction method for logging
        
    Returns:
        list: List of cleaned and filtered DataFrames
    """
    print(f"🧹 Cleaning {len(tables)} tables from {method_name}...")
    
    cleaned_tables = []
    for i, table in enumerate(tables):
        if table is not None and not table.empty:
            try:
                # Convert all data to string and handle NaN values
                table = table.fillna('')
                table = table.astype(str)
                
                # Clean up common extraction artifacts
                table = table.replace('nan', '')
                table = table.replace('None', '')
                
                # Remove completely empty rows and columns
                table = table.loc[:, (table != '').any(axis=0)]
                table = table.loc[(table != '').any(axis=1)]
                
                # Reset index
                table = table.reset_index(drop=True)
                
                # Clean individual cells
                for col in table.columns:
                    table[col] = table[col].astype(str)
                    table[col] = table[col].str.replace(r'[^\w\s\-\.\,\/\(\)\:]', ' ', regex=True)
                    table[col] = table[col].str.replace(r'\s+', ' ', regex=True)
                    table[col] = table[col].str.strip()
                
                # Basic filtering
                if not table.empty:
                    rows, cols = table.shape
                    
                    # Size check
                    if rows < 3 or cols < 2:
                        print(f"  ❌ Table {i+1}: Too small ({rows}×{cols})")
                        continue
                        
                    if rows > 200 or cols > 50:
                        print(f"  ❌ Table {i+1}: Too large ({rows}×{cols})")
                        continue
                    
                    # Content ratio check
                    non_empty_cells = (table != '').sum().sum()
                    total_cells = rows * cols
                    content_ratio = non_empty_cells / total_cells if total_cells > 0 else 0
                    
                    if content_ratio < 0.1:
                        print(f"  ❌ Table {i+1}: Low content ratio ({content_ratio:.2f})")
                        continue
                    
                    cleaned_tables.append(table)
                    print(f"  ✅ Table {i+1}: {rows}×{cols} (content: {content_ratio:.2f})")
                        
            except Exception as e:
                print(f"  ❌ Error processing table {i+1}: {e}")
                continue
    
    return cleaned_tables


def detect_pdf_type(pdf_path):
    """
    Detect if a PDF is primarily image-based or text-based using PDFplumber.
    Returns 'text' if PDF has searchable text, 'image' if it's primarily image-based.
    """
    try:
        import pdfplumber
        
        with pdfplumber.open(pdf_path) as pdf:
            text_content = ""
            pages_checked = 0
            
            # Check first few pages for text content
            for page in pdf.pages[:min(3, len(pdf.pages))]:
                page_text = page.extract_text()
                if page_text:
                    text_content += page_text
                pages_checked += 1
            
            # If we found substantial text content, it's likely text-based
            if len(text_content.strip()) > 100:  # At least 100 characters of text
                print(f"📄 PDF Type: TEXT-BASED (found {len(text_content)} characters)")
                return 'text'
            else:
                print(f"📄 PDF Type: IMAGE-BASED (found only {len(text_content)} characters)")
                return 'image'
                
    except ImportError:
        print("⚠️ PDFplumber not available - assuming text-based PDF")
        return 'text'
    except Exception as e:
        print(f"⚠️ Error detecting PDF type: {e} - assuming text-based PDF")
        return 'text'

def is_likely_bom_table(df):
    """
    Check if a DataFrame contains BOM-like data.
    Returns True if the table appears to be a Bill of Materials.
    """
    if df is None or df.empty:
        return False
    
    # BOM table detection constants
    HEADER_SYNONYMS = {
        'item':        ['ITEM', 'ITEM NO', 'ITEM NO.', 'ITEM NUMBER', 'NO', 'POS', 'POSITION'],
        'quantity':    ['QTY', 'QUANTITY', 'QUANT.', 'AMOUNT'],
        'description': ['DESCRIPTION', 'DESC', 'DETAILS', 'PART DESCRIPTION', 'PART NAME', 'ITEM DESCRIPTION', 'DEVICE'],
        'manufacturer': ['MANUFACTURER', 'MFG', 'MFR', 'MAKE', 'BRAND'],
        'part_number': ['PART NO', 'PART NUMBER', 'P/N', 'PN', 'MODEL NO', 'MODEL', 'MPN']
    }
    
    def has_any_synonym(header_cells, variants):
        """Check if any header cell contains variants of a field name."""
        for cell in header_cells:
            cell_upper = str(cell).upper().strip()
            if any(variant in cell_upper for variant in variants):
                return True
        return False
    
    try:
        # Get header row (first row)
        header_row = df.iloc[0].fillna('').astype(str)
        header_cells = [str(cell).upper().strip() for cell in header_row]
        
        # Check header fill ratio
        non_empty_headers = sum(1 for cell in header_cells if cell != '' and cell != 'NAN')
        header_fill_ratio = non_empty_headers / len(header_cells)
        
        if header_fill_ratio < 0.5:
            print(f"    ❌ BOM Check: Low header fill ratio ({header_fill_ratio:.2f})")
            return False
        
        # Check for required BOM column types
        required_fields = ['item', 'quantity']
        missing_fields = []
        
        for field in required_fields:
            if not has_any_synonym(header_cells, HEADER_SYNONYMS[field]):
                missing_fields.append(field)
        
        if missing_fields:
            print(f"    ❌ BOM Check: Missing required fields: {missing_fields}")
            return False
        
        # Additional scoring for table quality
        score = 0
        
        # Look for manufacturer/part number columns
        if has_any_synonym(header_cells, HEADER_SYNONYMS['manufacturer']):
            score += 2
        if has_any_synonym(header_cells, HEADER_SYNONYMS['part_number']):
            score += 2
        
        # Check for numeric patterns
        numeric_cols = 0
        for col in df.columns:
            col_data = df[col].astype(str).str.strip()
            numeric_count = sum(1 for val in col_data if val.replace('.', '').replace('-', '').isdigit())
            if numeric_count > len(col_data) * 0.3:
                numeric_cols += 1
        
        if numeric_cols >= 2:
            score += 2
        elif numeric_cols >= 1:
            score += 1
        
        # Check data density
        non_empty_cells = df.notna().sum().sum()
        total_cells = df.shape[0] * df.shape[1]
        data_density = non_empty_cells / total_cells
        
        if data_density < 0.3:
            score -= 2
        elif data_density > 0.6:
            score += 2
        elif data_density > 0.4:
            score += 1
        
        # Check for strong BOM indicators
        all_text = ' '.join(df.fillna('').astype(str).values.flatten()).upper()
        strong_bom_indicators = [
            'BILL OF MATERIAL', 'ITEM NO', 'MFG P/N', 'PROTON P/N', 'ALPHA WIRE',
            'HEYCO', 'SIEMENS', 'DELPHI', 'THOMAS BETTS', 'ALTECH', 'MURR',
            'PHOENIX', 'SQUARE D', 'EATON', 'CUSTOM FAB', 'SAGINAW', 'ASSEMBLY',
            'COMPONENT', 'HARDWARE', 'ELECTRICAL', 'MECHANICAL', 'FASTENER'
        ]
        
        strong_positive_score = sum(1 for indicator in strong_bom_indicators if indicator in all_text)
        score += strong_positive_score
        
        # Reject clearly non-BOM tables
        reject_keywords = [
            'PRINTED DRAWING', 'REFERENCE ONLY', 'DOCUMENT CONTROL', 'LATEST REVISION',
            'PROPERTY OF', 'DELIVERED ON', 'EXPRESS CONDITION', 'NOT TO BE DISCLOSED',
            'REVISIONS', 'ZONE', 'DESCRIPTION', 'DATE', 'APPROVED', 'RELEASE DATE'
        ]
        
        reject_score = sum(1 for keyword in reject_keywords if keyword in all_text)
        score -= reject_score * 2
        
        # Check for revision table patterns
        revision_patterns = ['REVISIONS', 'ZONE', 'REV', 'DESCRIPTION', 'DATE', 'APPROVED']
        revision_indicators = sum(1 for pattern in revision_patterns if pattern in all_text)
        
        if revision_indicators >= 4:
            print(f"    ❌ BOM Check: Appears to be revision table")
            return False
        
        threshold = 2
        result = score >= threshold
        
        print(f"    {'✅ BOM Check: PASS' if result else '❌ BOM Check: FAIL'} (score: {score}/{threshold})")
        return result
        
    except Exception as e:
        print(f"    ❌ BOM Check: Error during validation: {e}")
        return False

def extract_tables_from_pdf(pdf_path, pages):
    """
    Main extraction workflow that orchestrates the entire table extraction process.
    
    Workflow:
    1. Try tabula extraction first
    2. If tabula fails, check PDF type:
       - If image-based: go to OCR preprocessing
       - If text-based: try camelot fallback
    3. If camelot fails on text-based PDF: force OCR preprocessing
    4. After OCR: try tabula again, then camelot if needed
    5. If all methods fail: stop pipeline
    
    Args:
        pdf_path (str): Path to the PDF file
        pages (str): Page range to extract from (e.g., "1-3" or "all")
        
    Returns:
        list: List of extracted DataFrames that appear to be BOM tables
    """
    print(f"\n🚀 Starting main extraction workflow")
    print(f"📄 PDF: {pdf_path}")
    print(f"📄 Pages: {pages}")
    
    ocr_pdf_path = None
    final_tables = []
    
    try:
        # Step 1: Try tabula extraction first
        print(f"\n📊 STEP 1: Trying tabula extraction...")
        tabula_tables = extract_tables_with_tabula_method_impl(pdf_path, pages)
        
        if tabula_tables:
            # Check if any tables are likely BOM tables
            bom_tables = [table for table in tabula_tables if is_likely_bom_table(table)]
            
            if bom_tables:
                print(f"✅ STEP 1 SUCCESS: Found {len(bom_tables)} BOM tables with tabula")
                final_tables.extend(bom_tables)
                return final_tables
            else:
                print(f"❌ STEP 1 PARTIAL: Tabula found {len(tabula_tables)} tables but none appear to be BOM tables")
        else:
            print(f"❌ STEP 1 FAIL: Tabula found no tables")
        
        # Step 2: Tabula failed, check PDF type for next strategy
        print(f"\n🔍 STEP 2: Determining next extraction strategy...")
        pdf_type = detect_pdf_type(pdf_path)
        
        if pdf_type == 'image':
            print(f"📄 PDF is image-based - proceeding to OCR preprocessing")
            # Skip camelot for image-based PDFs, go straight to OCR
            ocr_pdf_path = process_pdf_with_ocr(pdf_path, force_ocr=True)
            
            if not ocr_pdf_path:
                print(f"❌ STEP 2 FAIL: OCR preprocessing failed")
                return []
                
        else:  # text-based PDF
            print(f"📄 PDF is text-based - trying camelot fallback")
            
            # Try camelot on original PDF
            camelot_tables = extract_tables_with_camelot_method(pdf_path, pages)
            
            if camelot_tables:
                bom_tables = [table for table in camelot_tables if is_likely_bom_table(table)]
                
                if bom_tables:
                    print(f"✅ STEP 2 SUCCESS: Found {len(bom_tables)} BOM tables with camelot")
                    final_tables.extend(bom_tables)
                    return final_tables
                else:
                    print(f"❌ STEP 2 PARTIAL: Camelot found {len(camelot_tables)} tables but none appear to be BOM tables")
            else:
                print(f"❌ STEP 2 FAIL: Camelot found no tables")
            
            # Camelot failed on text-based PDF, force OCR (might be vector images)
            print(f"📄 Text-based extraction failed - forcing OCR (may be vector images)")
            ocr_pdf_path = process_pdf_with_ocr(pdf_path, force_ocr=True)
            
            if not ocr_pdf_path:
                print(f"❌ STEP 2 FAIL: OCR preprocessing failed")
                return []
        
        # Step 3: Try extraction on OCR'd PDF
        print(f"\n🔄 STEP 3: Trying extraction on OCR'd PDF...")
        print(f"📄 OCR'd PDF: {ocr_pdf_path}")
        
        # Try tabula on OCR'd PDF first
        print(f"📊 STEP 3a: Trying tabula on OCR'd PDF...")
        ocr_tabula_tables = extract_tables_with_tabula_method_impl(ocr_pdf_path, pages)
        
        if ocr_tabula_tables:
            bom_tables = [table for table in ocr_tabula_tables if is_likely_bom_table(table)]
            
            if bom_tables:
                print(f"✅ STEP 3a SUCCESS: Found {len(bom_tables)} BOM tables with tabula on OCR'd PDF")
                final_tables.extend(bom_tables)
                return final_tables
            else:
                print(f"❌ STEP 3a PARTIAL: Tabula found {len(ocr_tabula_tables)} tables but none appear to be BOM tables")
        else:
            print(f"❌ STEP 3a FAIL: Tabula found no tables on OCR'd PDF")
        
        # Try camelot on OCR'd PDF
        print(f"🐪 STEP 3b: Trying camelot on OCR'd PDF...")
        ocr_camelot_tables = extract_tables_with_camelot_method(ocr_pdf_path, pages)
        
        if ocr_camelot_tables:
            bom_tables = [table for table in ocr_camelot_tables if is_likely_bom_table(table)]
            
            if bom_tables:
                print(f"✅ STEP 3b SUCCESS: Found {len(bom_tables)} BOM tables with camelot on OCR'd PDF")
                final_tables.extend(bom_tables)
                return final_tables
            else:
                print(f"❌ STEP 3b PARTIAL: Camelot found {len(ocr_camelot_tables)} tables but none appear to be BOM tables")
        else:
            print(f"❌ STEP 3b FAIL: Camelot found no tables on OCR'd PDF")
        
        # Step 4: All methods failed
        print(f"\n❌ STEP 4: All extraction methods failed")
        print(f"📊 Extraction pipeline complete - no BOM tables found")
        
        return []
        
    except Exception as e:
        print(f"❌ EXTRACTION ERROR: {e}")
        print(f"🔍 Error type: {type(e).__name__}")
        traceback.print_exc()
        return []
        
    finally:
        # Cleanup OCR temporary files
        if ocr_pdf_path and "bomination_ocr_" in str(ocr_pdf_path):
            try:
                cleanup_ocr_temp_files(ocr_pdf_path)
                print("🧹 Cleaned up OCR temporary files")
            except Exception as cleanup_error:
                print(f"⚠️ Warning: Could not cleanup OCR files: {cleanup_error}")

def clean_and_filter_tables(tables, method_name):
    """
    Clean extracted tables and filter for reasonable candidates.
    """
    print(f"🧹 Cleaning {len(tables)} tables from {method_name}...")
    
    cleaned_tables = []
    for i, table in enumerate(tables):
        if table is not None and not table.empty:
            try:
                # Convert all data to string and handle NaN values
                table = table.fillna('')
                table = table.astype(str)
                
                # Clean up common extraction artifacts
                table = table.replace('nan', '')
                table = table.replace('None', '')
                
                # Remove completely empty rows and columns
                table = table.loc[:, (table != '').any(axis=0)]
                table = table.loc[(table != '').any(axis=1)]
                
                # Reset index
                table = table.reset_index(drop=True)
                
                # Clean individual cells
                for col in table.columns:
                    table[col] = table[col].astype(str)
                    table[col] = table[col].str.replace(r'[^\w\s\-\.\,\/\(\)\:]', ' ', regex=True)
                    table[col] = table[col].str.replace(r'\s+', ' ', regex=True)
                    table[col] = table[col].str.strip()
                
                # Basic filtering
                if not table.empty:
                    rows, cols = table.shape
                    
                    # Size check
                    if rows < 3 or cols < 2:
                        print(f"  ❌ Table {i+1}: Too small ({rows}×{cols})")
                        continue
                    
                    if rows > 200 or cols > 30:
                        print(f"  ❌ Table {i+1}: Too large ({rows}×{cols})")
                        continue
                    
                    # Content ratio check
                    non_empty_cells = (table != '').sum().sum()
                    total_cells = rows * cols
                    content_ratio = non_empty_cells / total_cells if total_cells > 0 else 0
                    
                    if content_ratio < 0.1:
                        print(f"  ❌ Table {i+1}: Low content ratio ({content_ratio:.2f})")
                        continue
                    
                    cleaned_tables.append(table)
                    print(f"  ✅ Table {i+1}: {rows}×{cols} (content: {content_ratio:.2f})")
                        
            except Exception as e:
                print(f"  ❌ Error processing table {i+1}: {e}")
                continue
    
    return cleaned_tables

def extract_tables_with_enhanced_ocr_fallback(pdf_path, pages):
    """
    Try table extraction with enhanced OCR when automatic extraction fails.
    
    This function is designed for PDFs where:
    1. Automatic extraction finds tables but they're not useful (low text quality)
    2. The actual BoM content is image-based or has poor OCR quality
    3. Enhanced OCR preprocessing might improve table extraction
    
    Args:
        pdf_path (str): Path to the PDF file
        pages (str): Page range to extract from
        
    Returns:
        list: List of extracted DataFrames
    """
    print(f"\n🔍 Starting enhanced OCR fallback extraction")
    print(f"📄 PDF: {pdf_path}")
    print(f"📄 Pages: {pages}")
    
    # Try enhanced OCR preprocessing
    print("\n🔧 STEP 1: Applying table-optimized OCR preprocessing...")
    
    try:
        from pipeline.ocr_preprocessor import preprocess_pdf_for_table_extraction
        
        # Apply enhanced OCR
        success, ocr_pdf_path, error_msg = preprocess_pdf_for_table_extraction(
            pdf_path, 
            enhance_for_tables=True
        )
        
        if not success:
            print(f"❌ STEP 1 FAIL: Enhanced OCR preprocessing failed: {error_msg}")
            return []
            
        print(f"✅ STEP 1 SUCCESS: Enhanced OCR preprocessing completed")
        print(f"📁 OCR-enhanced PDF: {ocr_pdf_path}")
        
        # Try extraction on the OCR-enhanced PDF
        print("\n📊 STEP 2: Attempting table extraction on OCR-enhanced PDF...")
        
        # First try with tabula
        try:
            from pipeline.extract_bom_tab import extract_tables_with_tabula_method
            
            tabula_tables = extract_tables_with_tabula_method(
                ocr_pdf_path, 
                pages, 
                method='auto'
            )
            
            if tabula_tables:
                print(f"✅ STEP 2 SUCCESS: Found {len(tabula_tables)} tables with tabula on OCR-enhanced PDF")
                return tabula_tables
            else:
                print(f"❌ STEP 2 FAIL: No tables found with tabula on OCR-enhanced PDF")
                
        except Exception as e:
            print(f"❌ STEP 2 ERROR: Tabula extraction failed: {e}")
        
        # Try with camelot if tabula failed
        print("\n📊 STEP 3: Attempting camelot extraction on OCR-enhanced PDF...")
        
        try:
            from pipeline.extract_bom_cam import extract_tables_with_camelot_method
            
            camelot_tables = extract_tables_with_camelot_method(
                ocr_pdf_path, 
                pages
            )
            
            if camelot_tables:
                print(f"✅ STEP 3 SUCCESS: Found {len(camelot_tables)} tables with camelot on OCR-enhanced PDF")
                return camelot_tables
            else:
                print(f"❌ STEP 3 FAIL: No tables found with camelot on OCR-enhanced PDF")
                
        except Exception as e:
            print(f"❌ STEP 3 ERROR: Camelot extraction failed: {e}")
        
        # Clean up temporary OCR file
        try:
            if ocr_pdf_path != pdf_path and os.path.exists(ocr_pdf_path):
                os.remove(ocr_pdf_path)
                print(f"🧹 Cleaned up temporary OCR file: {ocr_pdf_path}")
        except Exception as e:
            print(f"⚠️ Could not clean up temporary file: {e}")
        
        print(f"❌ Enhanced OCR fallback extraction failed - no tables found")
        return []
        
    except Exception as e:
        print(f"❌ Enhanced OCR fallback extraction failed: {e}")
        return []


def extract_tables_with_roi_orchestration(pdf_path, pages):
    """
    Orchestrate ROI-based table extraction with proper fallback handling.
    
    This function coordinates between tabula and camelot for ROI-based extraction,
    maintaining proper separation of concerns where extract_bom_tab handles only
    tabula-specific operations and extract_main handles orchestration.
    
    Args:
        pdf_path (str): Path to the PDF file
        pages (str): Page range to extract from
        
    Returns:
        list: List of extracted DataFrames
    """
    print(f"\n🎯 Starting ROI-based extraction orchestration")
    print(f"� DEBUG: extract_tables_with_roi_orchestration() called")
    print(f"�📄 PDF: {pdf_path}")
    print(f"📄 Pages: {pages}")
    
    # Debug: Check ROI areas environment variable
    roi_areas_env = os.environ.get("BOM_ROI_AREAS")
    print(f"🐛 DEBUG: BOM_ROI_AREAS environment variable: {roi_areas_env}")
    
    # First, try tabula-only ROI extraction
    print("\n📊 STEP 1: Attempting tabula ROI extraction...")
    
    # Debug: Show which PDF file is being used
    print(f"🐛 DEBUG: PDF file for tabula ROI extraction: {pdf_path}")
    print(f"🐛 DEBUG: PDF file exists: {os.path.exists(pdf_path)}")
    
    # Check if this is an OCR-processed PDF
    from pathlib import Path
    is_ocr_pdf = "_ocr" in Path(pdf_path).stem
    print(f"🐛 DEBUG: Is OCR-processed PDF: {is_ocr_pdf}")
    
    if is_ocr_pdf:
        print(f"🔍 DEBUG: Using OCR-processed PDF - tabula ROI extraction will be skipped")
        print(f"🔍 DEBUG: Expected behavior: tabula will return empty, triggering Camelot fallback")
    else:
        print(f"🔍 DEBUG: Using original PDF - tabula ROI extraction will attempt normally")
    
    tabula_tables = extract_tables_with_roi_selection(pdf_path, pages)
    
    if tabula_tables:
        print(f"✅ STEP 1 SUCCESS: Tabula ROI extraction found {len(tabula_tables)} tables")
        return tabula_tables
    
    print(f"❌ STEP 1 FAIL: Tabula ROI extraction found no tables")
    
    # If tabula failed, try camelot with ROI areas
    print("\n📊 STEP 2: Attempting camelot ROI extraction...")
    
    # Debug: Show which PDF file Camelot will use
    print(f"🐛 DEBUG: PDF file for camelot ROI extraction: {pdf_path}")
    print(f"🐛 DEBUG: PDF file exists: {os.path.exists(pdf_path)}")
    print(f"🐛 DEBUG: Is OCR-processed PDF: {is_ocr_pdf}")
    
    if is_ocr_pdf:
        print(f"🔍 DEBUG: Camelot will use OCR-processed PDF - should handle OCR'd tables better")
    else:
        print(f"🔍 DEBUG: Camelot will use original PDF")
    
    roi_areas = os.environ.get("BOM_ROI_AREAS")
    
    if not roi_areas:
        print("❌ STEP 2 FAIL: No ROI areas available for camelot fallback")
        return []
    
    try:
        import json
        roi_areas = json.loads(roi_areas)
        
        all_tables = []
        
        for page_num, area in roi_areas.items():
            print(f"\n📊 Extracting from page {page_num} using Camelot ROI area: {area}")
            
            try:
                # Convert area coordinates for Camelot
                # Tabula area: [top, left, bottom, right] in points
                # Camelot ROI function expects: [left, top, right, bottom] in points
                camelot_area = [area[1], area[0], area[3], area[2]]
                
                print(f"🐛 DEBUG: Converting ROI area for Camelot:")
                print(f"  Original (tabula): {area}")
                print(f"  Converted (camelot): {camelot_area}")
                print(f"  📐 Coordinate System Info:")
                print(f"    - Tabula format: [top, left, bottom, right] (origin: top-left)")
                print(f"    - Camelot format: [x1, y1, x2, y2] where (x1,y1)=top-left, (x2,y2)=bottom-right")
                print(f"    - PDF coordinate space: (0,0) = bottom-left corner of page")
                print(f"  🔄 Conversion: tabula[top,left,bottom,right] → camelot[left,top,right,bottom]")
                
                # Visual debug: Draw the ROI area on PDF to verify coordinates
                try:
                    debug_image_path = visualize_camelot_roi_on_pdf(
                        pdf_path, int(page_num), camelot_area, area
                    )
                    if debug_image_path:
                        print(f"  📷 Debug ROI visualization saved: {debug_image_path}")
                except Exception as viz_error:
                    print(f"  ⚠️ Could not create ROI visualization: {viz_error}")
                
                # Use the ROI-specific Camelot function
                from pipeline.extract_bom_cam import extract_tables_with_camelot_roi
                camelot_tables = extract_tables_with_camelot_roi(
                    pdf_path, 
                    str(page_num),
                    roi_areas=[camelot_area]
                )
                
                if camelot_tables:
                    for i, table in enumerate(camelot_tables):
                        if not table.empty and table.shape[0] >= 2 and table.shape[1] >= 2:
                            print(f"    ✅ Extracted table {i+1}: {table.shape[0]}×{table.shape[1]} (Camelot ROI)")
                            all_tables.append(table)
                        else:
                            print(f"    ❌ Camelot table {i+1} too small: {table.shape[0]}×{table.shape[1]}")
                else:
                    print(f"    ❌ No Camelot tables from page {page_num}")
                    
            except Exception as e:
                print(f"    ❌ Camelot extraction failed for page {page_num}: {e}")
                import traceback
                traceback.print_exc()
                continue
        
        if all_tables:
            print(f"✅ STEP 2 SUCCESS: Camelot ROI extraction found {len(all_tables)} tables")
            return all_tables
        else:
            print(f"❌ STEP 2 FAIL: Camelot ROI extraction found no tables")
            return []
            
    except Exception as e:
        print(f"❌ STEP 2 ERROR: Camelot ROI orchestration failed: {e}")
        return []

def extract_tables_from_pdf_auto(pdf_path, pages="all", extraction_method="auto"):
    """
    Main entry point for table extraction from PDF files.
    
    This function orchestrates the entire extraction workflow:
    1. Tabula extraction
    2. PDF type detection and fallback strategy
    3. OCR preprocessing if needed
    4. Table validation and filtering
    
    Args:
        pdf_path (str): Path to the PDF file
        pages (str): Page range to extract from (e.g., "1-3", "all")
        extraction_method (str): "auto", "tabula", "camelot", "ocr", or "roi"
        
    Returns:
        list: List of extracted and validated DataFrames
    """
    print(f"\n🚀 Starting BoM table extraction")
    print(f"📄 PDF: {pdf_path}")
    print(f"📄 Pages: {pages}")
    print(f"🔧 Method: {extraction_method}")
    
    try:
        # Validate input
        if not os.path.exists(pdf_path):
            print(f"❌ PDF file not found: {pdf_path}")
            return []
        
        # Handle special extraction methods
        if extraction_method == "roi":
            from pipeline.extract_bom_tab import extract_tables_with_roi_selection
            return extract_tables_with_roi_selection(pdf_path, pages)
        elif extraction_method == "tabula":
            from pipeline.extract_bom_tab import extract_tables_with_tabula_method_impl
            tables = extract_tables_with_tabula_method_impl(pdf_path, pages)
            return clean_and_filter_tables(tables, "tabula")
        elif extraction_method == "camelot":
            from pipeline.extract_bom_cam import extract_tables_with_camelot_method
            tables = extract_tables_with_camelot_method(pdf_path, pages)
            return clean_and_filter_tables(tables, "camelot")
        elif extraction_method == "ocr":
            from pipeline.ocr_preprocessor import process_pdf_with_ocr
            ocr_pdf_path = process_pdf_with_ocr(pdf_path, force_ocr=True)
            if ocr_pdf_path:
                # Try tabula on OCR'd PDF first
                from pipeline.extract_bom_tab import extract_tables_with_tabula_method_impl
                tables = extract_tables_with_tabula_method_impl(ocr_pdf_path, pages)
                return clean_and_filter_tables(tables, "ocr-tabula")
            else:
                print("❌ OCR preprocessing failed")
                return []
        else:  # auto method
            # Use local extraction methods in sequence
            # First try tabula
            from pipeline.extract_bom_tab import extract_tables_with_tabula_method_impl
            tables = extract_tables_with_tabula_method_impl(pdf_path, pages)
            cleaned_tables = clean_and_filter_tables(tables, "auto-tabula")
            
            if cleaned_tables:
                # Check if tables are of good quality
                good_quality_tables = []
                for table in cleaned_tables:
                    # Check table quality based on content density and readability
                    if is_table_good_quality(table):
                        good_quality_tables.append(table)
                
                if good_quality_tables:
                    print(f"✅ Found {len(good_quality_tables)} good quality tables with tabula")
                    return good_quality_tables
                else:
                    print(f"⚠️ Found {len(cleaned_tables)} tables with tabula, but quality is poor")
                    print("🔍 Consider using ROI mode for better extraction")
                    return cleaned_tables
            
            # If tabula fails completely, delegate to main orchestrator
            print("� Tabula found no tables - delegating to main orchestrator")
            return []
            
    except Exception as e:
        print(f"❌ Extraction failed: {e}")
        traceback.print_exc()
        return []

def process_and_format_tables(tables, customer_name=""):
    """
    Process extracted tables and apply customer-specific formatting.
    
    Args:
        tables (list): List of DataFrames to process
        customer_name (str): Name of customer for specific formatting
        
    Returns:
        list: List of processed and formatted DataFrames
    """
    if not tables:
        return []
    
    print(f"\n🔧 Processing {len(tables)} tables...")
    
    processed_tables = []
    
    for i, table in enumerate(tables):
        try:
            print(f"📊 Processing table {i+1}...")
            
            # Clean and validate table
            cleaned_table = clean_and_filter_tables([table], f"table_{i+1}")
            
            if cleaned_table:
                table = cleaned_table[0]
                
                # Apply customer formatting if specified
                if customer_name:
                    print(f"🎨 Applying {customer_name} formatting...")
                    from omni_cust.customer_formatters import apply_customer_formatter
                    formatted_table = apply_customer_formatter(table, customer_name)
                    if formatted_table is not None:
                        table = formatted_table
                
                processed_tables.append(table)
                print(f"    ✅ Table {i+1} processed successfully")
            else:
                print(f"    ❌ Table {i+1} failed validation")
                
        except Exception as e:
            print(f"    ❌ Error processing table {i+1}: {e}")
            continue
    
    print(f"✅ Processing complete: {len(processed_tables)} tables ready")
    return processed_tables

def save_tables_to_excel(tables, output_path):
    """
    Save processed tables to an Excel file.
    
    Args:
        tables (list): List of DataFrames to save
        output_path (str): Path to save the Excel file
        
    Returns:
        bool: True if successful, False otherwise
    """
    if not tables:
        print("❌ No tables to save")
        return False
    
    try:
        import pandas as pd
        print(f"💾 Saving {len(tables)} tables to Excel...")
        print(f"📁 Output path: {output_path}")
        
        # Ensure output directory exists
        from pathlib import Path
        output_dir = Path(output_path).parent
        output_dir.mkdir(parents=True, exist_ok=True)
        
        # Save to Excel with multiple sheets
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            for i, table in enumerate(tables):
                sheet_name = f"Table_{i+1}"
                table.to_excel(writer, sheet_name=sheet_name, index=False)
                print(f"    ✅ Saved table {i+1} to sheet '{sheet_name}'")
        
        print(f"✅ Excel file saved successfully: {output_path}")
        return True
        
    except Exception as e:
        print(f"❌ Error saving Excel file: {e}")
        import traceback
        traceback.print_exc()
        return False

def merge_tables_and_export(tables, output_path, sheet_name="Combined_BoM", company=""):
    """
    Merge multiple tables and export to a single Excel sheet.
    
    Args:
        tables (list): List of DataFrames to merge
        output_path (str): Path to save the Excel file
        sheet_name (str): Name of the Excel sheet
        company (str): Company name for formatting
        
    Returns:
        bool: True if successful, False otherwise
    """
    if not tables:
        print("❌ No tables to merge")
        return False
    
    try:
        import pandas as pd
        
        print(f"🔗 Merging {len(tables)} tables...")
        if company:
            print(f"🎨 Will apply {company} customer formatting...")
        
        # Apply customer formatting to each table before merging
        formatted_tables = []
        for i, table in enumerate(tables):
            print(f"📊 Processing table {i+1} for merging...")
            
            # Apply customer formatting if specified
            if company:
                print(f"🎨 Applying {company} formatting to table {i+1}...")
                try:
                    from omni_cust.customer_formatters import apply_customer_formatter
                    formatted_table = apply_customer_formatter(table, company)
                    if formatted_table is not None and not formatted_table.empty:
                        formatted_tables.append(formatted_table)
                        print(f"    ✅ Table {i+1} formatted successfully")
                    else:
                        print(f"    ⚠️ Table {i+1} formatting returned empty, using original")
                        formatted_tables.append(table)
                except Exception as e:
                    print(f"    ❌ Error formatting table {i+1}: {e}")
                    print(f"    ⚠️ Using original table without formatting")
                    formatted_tables.append(table)
            else:
                formatted_tables.append(table)
        
        # Simple concatenation for now - could be enhanced with more sophisticated merging
        merged_table = pd.concat(formatted_tables, ignore_index=True)
        
        # Clean merged table
        merged_table = merged_table.fillna('')
        merged_table = merged_table.replace('nan', '')
        
        # Remove completely empty rows
        merged_table = merged_table.loc[(merged_table != '').any(axis=1)]
        
        print(f"✅ Merged table created: {merged_table.shape[0]}×{merged_table.shape[1]}")
        print(f"📊 Final table columns: {merged_table.columns.tolist()}")
        
        # Save to Excel
        print(f"💾 Saving merged table to Excel...")
        
        # Ensure output directory exists
        from pathlib import Path
        output_dir = Path(output_path).parent
        output_dir.mkdir(parents=True, exist_ok=True)
        
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            merged_table.to_excel(writer, sheet_name=sheet_name, index=False)
        
        print(f"✅ Merged table saved successfully: {output_path}")
        return True
        
    except Exception as e:
        print(f"❌ Error merging and saving tables: {e}")
        import traceback
        traceback.print_exc()
        return False

def format_table_as_text(table):
    """Format a pandas DataFrame as readable text with proper spacing."""
    if table.empty:
        return "Empty table"
    
    # Get column headers
    headers = [f"Col_{j}" for j in range(len(table.columns))]
    
    # Calculate column widths
    col_widths = []
    for j in range(len(table.columns)):
        max_width = max(len(str(headers[j])), 
                       max(len(str(cell)) for cell in table.iloc[:, j]))
        col_widths.append(min(max_width, 30))  # Cap at 30 characters
    
    # Format the table text
    text_lines = []
    
    # Header row
    header_line = " | ".join(headers[j].ljust(col_widths[j]) for j in range(len(headers)))
    text_lines.append(header_line)
    text_lines.append("-" * len(header_line))
    
    # Data rows
    for i in range(len(table)):
        row_line = " | ".join(str(table.iloc[i, j]).ljust(col_widths[j]) for j in range(len(table.columns)))
        text_lines.append(row_line)
    
    return "\n".join(text_lines)

# Constants for backward compatibility
MIN_ROWS, MIN_COLS = 3, 3  # Minimum table size to consider
HEADER_SYNONYMS = {
    'item':        ['ITEM', 'ITEM NO', 'ITEM NO.', 'ITEM NUMBER', 'NO', 'POS', 'POSITION'],
    'quantity':    ['QTY', 'QUANTITY', 'QUANT.', 'AMOUNT'],
    'description': ['DESCRIPTION', 'DESC', 'DETAILS', 'PART DESCRIPTION', 'PART NAME', 'ITEM DESCRIPTION', 'DEVICE'],
    'manufacturer': ['MANUFACTURER', 'MFG', 'MFR', 'MAKE', 'BRAND'],
    'part_number': ['PART NO', 'PART NUMBER', 'P/N', 'PN', 'MODEL NO', 'MODEL', 'MPN']
}

def clean_table_headers(table):
    """Clean and normalize table headers."""
    if table.empty:
        return table
    
    # Make a copy
    cleaned = table.copy()
    
    # Clean the first row (assumed to be headers)
    if len(cleaned) > 0:
        first_row = cleaned.iloc[0]
        cleaned_headers = []
        
        for header in first_row:
            header_str = str(header).strip().upper()
            # Remove common artifacts
            header_str = header_str.replace('\n', ' ').replace('\r', ' ')
            header_str = ' '.join(header_str.split())  # Normalize whitespace
            cleaned_headers.append(header_str)
        
        # Update the first row
        cleaned.iloc[0] = cleaned_headers
    
    return cleaned

def has_any_synonym(header_cells, variants):
    """Check if any header cell contains variants of a field name."""
    for cell in header_cells:
        cell_upper = str(cell).upper().strip()
        if any(variant in cell_upper for variant in variants):
            return True
    return False

def run_main_extraction_workflow():
    """
    Main execution function when script is run directly.
    This function handles the environment variables and coordinates the extraction workflow.
    """
    print("🚀 Starting BoMination table extraction...")
    print(f"🐛 DEBUG: run_main_extraction_workflow() called")
    
    # Get configuration from environment variables
    pdf_path = os.environ.get("BOM_PDF_PATH")
    pages = os.environ.get("BOM_PAGE_RANGE", "all")
    company = os.environ.get("BOM_COMPANY", "")
    output_directory = os.environ.get("BOM_OUTPUT_DIRECTORY", "")
    tabula_mode = os.environ.get("BOM_TABULA_MODE", "balanced")
    use_roi = os.environ.get("BOM_USE_ROI", "false").lower() == "true"
    
    print(f"🐛 DEBUG: Environment variables:")
    print(f"  BOM_PDF_PATH: {pdf_path}")
    print(f"  BOM_PAGE_RANGE: {pages}")
    print(f"  BOM_COMPANY: {company}")
    print(f"  BOM_OUTPUT_DIRECTORY: {output_directory}")
    print(f"  BOM_TABULA_MODE: {tabula_mode}")
    print(f"  BOM_USE_ROI: {os.environ.get('BOM_USE_ROI', 'NOT SET')} -> {use_roi}")
    
    # Check if ROI areas are set
    roi_areas = os.environ.get("BOM_ROI_AREAS", "NOT SET")
    print(f"  BOM_ROI_AREAS: {roi_areas}")
    
    # Validate inputs
    if not pdf_path:
        print("❌ No PDF path provided")
        import sys
        sys.exit(1)
    
    if not os.path.exists(pdf_path):
        print(f"❌ PDF file not found: {pdf_path}")
        import sys
        sys.exit(1)
    
    print(f"📄 PDF: {pdf_path}")
    print(f"📄 Pages: {pages}")
    print(f"🏢 Company: {company}")
    print(f"📁 Output directory: {output_directory}")
    print(f"🔧 Tabula mode: {tabula_mode}")
    print(f"📍 Use ROI: {use_roi}")
    
    try:
        # Determine extraction method based on ROI setting
        if use_roi:
            print("📍 Using ROI-based extraction...")
            # Use dedicated ROI extraction orchestration
            extracted_tables = extract_tables_with_roi_orchestration(pdf_path, pages)
        else:
            print("🔄 Using automatic extraction workflow...")
            # Use automatic extraction workflow
            extracted_tables = extract_tables_from_pdf_auto(pdf_path, pages, "auto")
        
        if not extracted_tables:
            print("❌ No tables extracted from PDF")
            import sys
            sys.exit(1)
        
        print(f"✅ Extracted {len(extracted_tables)} tables")
        
        # Show table selection interface
        if len(extracted_tables) > 1:
            print("📋 Multiple tables found - showing selection interface...")
            from gui.table_selector import show_table_selector
            selected_tables = show_table_selector(extracted_tables)
            
            if not selected_tables:
                print("❌ No tables selected by user")
                import sys
                sys.exit(1)
        else:
            print("📋 Single table found - using automatically...")
            selected_tables = extracted_tables
        
        # Process and format tables
        processed_tables = process_and_format_tables(selected_tables, company)
        
        if not processed_tables:
            print("❌ No tables passed processing")
            import sys
            sys.exit(1)
        
        # Generate output path
        from pathlib import Path
        pdf_dir = Path(pdf_path).parent
        pdf_name = Path(pdf_path).stem
        
        if output_directory:
            output_dir = Path(output_directory)
            output_dir.mkdir(parents=True, exist_ok=True)
            merged_path = output_dir / f"{pdf_name}_merged.xlsx"
        else:
            merged_path = pdf_dir / f"{pdf_name}_merged.xlsx"
        
        # Save merged tables
        success = merge_tables_and_export(processed_tables, str(merged_path), "Combined_BoM", company)
        
        if success:
            print(f"✅ Extraction completed successfully!")
            print(f"📁 Output file: {merged_path}")
        else:
            print("❌ Failed to save merged tables")
            import sys
            sys.exit(1)
            
    except Exception as e:
        print(f"❌ Extraction failed: {e}")
        import traceback
        traceback.print_exc()
        import sys
        sys.exit(1)

# Legacy function for backward compatibility
def extract_tables_with_tabula(pdf_path, pages):
    """Legacy function - redirects to main extraction workflow."""
    return extract_tables_from_pdf(pdf_path, pages)
