"""
Camelot-specific table extraction module for BoMination.
Handles table extraction from PDFs using the camelot-py library.
"""

import os
import sys
import pandas as pd
import logging
import traceback
from pathlib import Path

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')
logger = logging.getLogger(__name__)

def extract_tables_with_camelot_method(pdf_path, pages, table_areas=None):
    """
    Extract tables using Camelot with both lattice and stream methods.
    Works well with both text-based and image-based PDFs.
    
    Args:
        pdf_path (str): Path to the PDF file
        pages (str): Page range to extract from (e.g., "1-3" or "all")
        table_areas (list, optional): List of table areas in format [top, left, bottom, right]
                                    If provided, will restrict extraction to these areas only
        
    Returns:
        list: List of extracted DataFrames
    """
    print("🐪 Attempting table extraction with Camelot...")
    
    try:
        import camelot
        print(f"🐪 Camelot available - version check...")
    except ImportError:
        print("❌ Camelot not available - install with: pip install camelot-py[cv]")
        return []
    
    try:
        tables = []
        
        # Convert pages parameter to string format camelot expects
        if isinstance(pages, list):
            pages_str = ','.join(map(str, pages))
        else:
            pages_str = str(pages)
        
        print(f"🐪 Extracting from pages: {pages_str}")
        
        # Log ROI area usage
        if table_areas:
            print(f"🎯 Using ROI areas: {table_areas}")
        else:
            print(f"🌐 No ROI areas specified - extracting from entire page")
        
        # Method 1: Lattice method (best for tables with clear borders)
        print("🐪 Trying lattice method...")
        try:
            lattice_tables = camelot.read_pdf(
                pdf_path, 
                pages=pages_str, 
                flavor='lattice',
                split_text=True,
                flag_size=True,
                strip_text='\n',
                line_scale=40,
                copy_text=['h', 'v'],
                shift_text=['l', 't'],
                process_background=True,
                table_areas=table_areas  # Use ROI areas if provided
            )
            
            if lattice_tables and len(lattice_tables) > 0:
                print(f"    ✅ Lattice method found {len(lattice_tables)} tables")
                
                for i, camelot_table in enumerate(lattice_tables):
                    df = camelot_table.df
                    accuracy = camelot_table.accuracy
                    
                    print(f"    Table {i+1}: {df.shape[0]}×{df.shape[1]} (accuracy: {accuracy:.1f}%)")
                    
                    # Filter by accuracy and size
                    if not df.empty and df.shape[0] >= 3 and df.shape[1] >= 3 and accuracy > 50:
                        print(f"        ✅ Accepted (lattice)")
                        tables.append(df)
                    else:
                        print(f"        ❌ Rejected - size: {df.shape}, accuracy: {accuracy:.1f}%")
            else:
                print("    ❌ Lattice method found no tables")
                
        except Exception as e:
            print(f"    ❌ Lattice method failed: {e}")
        
        # Method 2: Stream method (good for tables without clear borders)
        if not tables:
            print("🐪 Trying stream method...")
            try:
                stream_tables = camelot.read_pdf(
                    pdf_path, 
                    pages=pages_str, 
                    flavor='stream',
                    split_text=True,
                    flag_size=True,
                    strip_text='\n',
                    edge_tol=500,
                    row_tol=10,
                    column_tol=0,
                    table_areas=table_areas  # Use ROI areas if provided
                )
                
                if stream_tables and len(stream_tables) > 0:
                    print(f"    ✅ Stream method found {len(stream_tables)} tables")
                    
                    for i, camelot_table in enumerate(stream_tables):
                        df = camelot_table.df
                        accuracy = camelot_table.accuracy
                        
                        print(f"    Table {i+1}: {df.shape[0]}×{df.shape[1]} (accuracy: {accuracy:.1f}%)")
                        
                        # More lenient criteria for stream method
                        if not df.empty and df.shape[0] >= 3 and df.shape[1] >= 3 and accuracy > 30:
                            print(f"        ✅ Accepted (stream)")
                            tables.append(df)
                        else:
                            print(f"        ❌ Rejected - size: {df.shape}, accuracy: {accuracy:.1f}%")
                else:
                    print("    ❌ Stream method found no tables")
                    
            except Exception as e:
                print(f"    ❌ Stream method failed: {e}")
        
        # Method 3: Lattice with lenient settings
        if not tables:
            print("🐪 Trying lattice with lenient settings...")
            try:
                lenient_tables = camelot.read_pdf(
                    pdf_path, 
                    pages=pages_str, 
                    flavor='lattice',
                    split_text=True,
                    flag_size=True,
                    strip_text='\n',
                    line_scale=20,  # More lenient line detection
                    copy_text=['h', 'v'],
                    shift_text=['l', 't'],
                    process_background=True,
                    table_areas=table_areas,  # Use ROI areas if provided
                    columns=None
                )
                
                if lenient_tables and len(lenient_tables) > 0:
                    print(f"    ✅ Lenient lattice found {len(lenient_tables)} tables")
                    
                    for i, camelot_table in enumerate(lenient_tables):
                        df = camelot_table.df
                        accuracy = camelot_table.accuracy
                        
                        print(f"    Table {i+1}: {df.shape[0]}×{df.shape[1]} (accuracy: {accuracy:.1f}%)")
                        
                        # Even more lenient criteria
                        if not df.empty and df.shape[0] >= 2 and df.shape[1] >= 2 and accuracy > 20:
                            print(f"        ✅ Accepted (lenient)")
                            tables.append(df)
                        else:
                            print(f"        ❌ Rejected - size: {df.shape}, accuracy: {accuracy:.1f}%")
                else:
                    print("    ❌ Lenient lattice found no tables")
                    
            except Exception as e:
                print(f"    ❌ Lenient lattice failed: {e}")
        
        # Method 4: Stream with very lenient settings
        if not tables:
            print("🐪 Trying stream with very lenient settings...")
            try:
                very_lenient_tables = camelot.read_pdf(
                    pdf_path, 
                    pages=pages_str, 
                    flavor='stream',
                    split_text=True,
                    flag_size=True,
                    strip_text='\n',
                    edge_tol=1000,  # Very lenient edge tolerance
                    row_tol=20,     # More lenient row tolerance
                    column_tol=10   # More lenient column tolerance
                )
                
                if very_lenient_tables and len(very_lenient_tables) > 0:
                    print(f"    ✅ Very lenient stream found {len(very_lenient_tables)} tables")
                    
                    for i, camelot_table in enumerate(very_lenient_tables):
                        df = camelot_table.df
                        accuracy = camelot_table.accuracy
                        
                        print(f"    Table {i+1}: {df.shape[0]}×{df.shape[1]} (accuracy: {accuracy:.1f}%)")
                        
                        # Very lenient criteria for last resort
                        if not df.empty and df.shape[0] >= 2 and df.shape[1] >= 2 and accuracy > 10:
                            print(f"        ✅ Accepted (very lenient)")
                            tables.append(df)
                        else:
                            print(f"        ❌ Rejected - size: {df.shape}, accuracy: {accuracy:.1f}%")
                else:
                    print("    ❌ Very lenient stream found no tables")
                    
            except Exception as e:
                print(f"    ❌ Very lenient stream failed: {e}")
        
        print(f"🐪 Camelot extraction completed: {len(tables)} tables found")
        
        # Show summary of extracted tables
        if tables:
            print("📋 Camelot extraction summary:")
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
        print(f"❌ Camelot extraction failed: {e}")
        print(f"🔍 Error type: {type(e).__name__}")
        traceback.print_exc()
        return []

def validate_camelot_table(table, accuracy=0):
    """
    Validate if a table extracted by camelot is worth processing.
    
    Args:
        table (DataFrame): Table to validate
        accuracy (float): Camelot accuracy score
        
    Returns:
        bool: True if table passes validation
    """
    if table is None or table.empty:
        return False
    
    rows, cols = table.shape
    
    # Basic size check
    if rows < 3 or cols < 2:
        return False
    
    # Check accuracy if provided
    if accuracy > 0 and accuracy < 20:
        return False
    
    # Check content density
    non_empty_cells = table.notna().sum().sum()
    total_cells = rows * cols
    density = non_empty_cells / total_cells
    
    if density < 0.1:  # Less than 10% content
        return False
    
    # Check for reasonable cell content lengths
    all_text = ' '.join(table.fillna('').astype(str).values.flatten())
    avg_cell_length = len(all_text) / total_cells
    
    if avg_cell_length > 500:  # Cells too long (poor extraction)
        return False
    
    return True

def clean_camelot_table(table):
    """
    Clean and normalize a table extracted by camelot.
    
    Args:
        table (DataFrame): Table to clean
        
    Returns:
        DataFrame: Cleaned table
    """
    if table is None or table.empty:
        return table
    
    # Make a copy to avoid modifying original
    cleaned = table.copy()
    
    # Fill NaN values with empty strings
    cleaned = cleaned.fillna('')
    
    # Convert to string and clean
    cleaned = cleaned.astype(str)
    cleaned = cleaned.replace('nan', '')
    cleaned = cleaned.replace('None', '')
    
    # Clean individual cells
    for col in cleaned.columns:
        cleaned[col] = cleaned[col].str.strip()
        cleaned[col] = cleaned[col].str.replace(r'\s+', ' ', regex=True)
        cleaned[col] = cleaned[col].str.replace(r'[^\w\s\-\.\,\/\(\)\:]', ' ', regex=True)
        cleaned[col] = cleaned[col].str.strip()
    
    # Remove completely empty rows and columns
    cleaned = cleaned.loc[:, (cleaned != '').any(axis=0)]
    cleaned = cleaned.loc[(cleaned != '').any(axis=1)]
    
    # Reset index
    cleaned = cleaned.reset_index(drop=True)
    
    return cleaned

def extract_tables_with_camelot_roi(pdf_path, pages, roi_areas=None):
    """
    Extract tables using Camelot with specific ROI areas.
    
    Args:
        pdf_path (str): Path to PDF file
        pages (str): Page range to extract from
        roi_areas (list): List of ROI areas in format [left, top, right, bottom]
        
    Returns:
        list: List of extracted DataFrames
    """
    if not roi_areas:
        return extract_tables_with_camelot_method(pdf_path, pages)
    
    print(f"🐪 Attempting Camelot extraction with ROI areas...")
    
    try:
        import camelot
    except ImportError:
        print("❌ Camelot not available - install with: pip install camelot-py[cv]")
        return []
    
    try:
        tables = []
        
        # Convert pages parameter
        if isinstance(pages, list):
            pages_str = ','.join(map(str, pages))
        else:
            pages_str = str(pages)
        
        # Convert ROI areas to camelot format
        # ROI picker uses matplotlib coordinates: (0,0) = top-left, Y increases downward
        # PDF uses PDF coordinate space: (0,0) = bottom-left, Y increases upward
        # We need to flip the Y coordinates
        
        table_areas = []
        
        # Get the actual page dimensions for coordinate conversion
        try:
            import fitz  # PyMuPDF
            doc = fitz.open(pdf_path)
            page = doc[0]  # Get first page
            page_height = page.rect.height
            page_width = page.rect.width
            doc.close()
            print(f"📏 Page dimensions: {page_width} x {page_height} points")
        except ImportError:
            print("⚠️ PyMuPDF not available, using standard page height")
            page_height = 792  # Standard letter size page height in points
        except Exception as e:
            print(f"⚠️ Could not get page dimensions: {e}, using standard page height")
            page_height = 792
        
        for area in roi_areas:
            # area format from ROI picker: [top, left, bottom, right] in matplotlib coordinates
            if len(area) == 4:
                # Convert from matplotlib to PDF coordinates
                # matplotlib: top=area[0], left=area[1], bottom=area[2], right=area[3]
                # PDF: we need to flip Y coordinates
                pdf_left = area[1]
                pdf_top = page_height - area[2]    # flip: bottom becomes top
                pdf_right = area[3]
                pdf_bottom = page_height - area[0]  # flip: top becomes bottom
                
                # Add small padding to ensure we capture the target table
                # Sometimes ROI selection is slightly off due to zoom/precision
                padding = 5  # 5 points padding
                pdf_left = max(0, pdf_left - padding)
                pdf_top = max(0, pdf_top - padding)
                pdf_right = min(page_width, pdf_right + padding)
                pdf_bottom = min(page_height, pdf_bottom + padding)
                
                # Camelot format: "left,top,right,bottom" in PDF coordinates
                camelot_area = f"{pdf_left},{pdf_top},{pdf_right},{pdf_bottom}"
                table_areas.append(camelot_area)
                
                print(f"🔧 ROI conversion (with Y-flip and padding):")
                print(f"   Matplotlib: [top={area[0]}, left={area[1]}, bottom={area[2]}, right={area[3]}]")
                print(f"   PDF coords: left={pdf_left}, top={pdf_top}, right={pdf_right}, bottom={pdf_bottom}")
                print(f"   Camelot: {camelot_area}")
                print(f"   Applied padding: {padding} points")
                
                # Visual debug: Create ROI visualization
                try:
                    from pipeline.extract_main import visualize_camelot_roi_on_pdf
                    camelot_coords = [pdf_left, pdf_top, pdf_right, pdf_bottom]
                    original_coords = [area[0], area[1], area[2], area[3]]  # [top, left, bottom, right]
                    
                    # Extract page number from pages parameter
                    try:
                        if isinstance(pages, str) and pages.isdigit():
                            page_num = int(pages)
                        elif isinstance(pages, int):
                            page_num = pages
                        else:
                            page_num = 1  # Default to page 1
                    except:
                        page_num = 1
                    
                    debug_image_path = visualize_camelot_roi_on_pdf(
                        pdf_path, page_num, camelot_coords, original_coords
                    )
                    if debug_image_path:
                        print(f"   📷 Debug ROI visualization saved: {debug_image_path}")
                    else:
                        print(f"   ⚠️ Debug visualization failed (no image created)")
                except Exception as viz_error:
                    print(f"   ⚠️ Could not create ROI visualization: {viz_error}")
        
        print(f"🐪 Using ROI areas: {table_areas}")
        
        # Try lattice method with ROI areas
        try:
            print(f"🔧 Calling camelot.read_pdf with table_areas={table_areas}")
            roi_tables = camelot.read_pdf(
                pdf_path,
                pages=pages_str,
                flavor='lattice',
                table_areas=table_areas,
                split_text=True,
                flag_size=True,
                strip_text='\n'
            )
            
            print(f"🔧 Camelot returned {len(roi_tables) if roi_tables else 0} tables")
            
            if roi_tables and len(roi_tables) > 0:
                print(f"    ✅ ROI lattice found {len(roi_tables)} tables")
                
                for i, camelot_table in enumerate(roi_tables):
                    df = camelot_table.df
                    accuracy = camelot_table.accuracy
                    
                    print(f"    Table {i+1}: {df.shape[0]}×{df.shape[1]} (accuracy: {accuracy:.1f}%)")
                    
                    # Show more detailed content to help debug
                    print(f"    Table {i+1} full content preview:")
                    print(f"    Row 0: {df.iloc[0].tolist() if len(df) > 0 else 'No data'}")
                    if len(df) > 1:
                        print(f"    Row 1: {df.iloc[1].tolist()}")
                    if len(df) > 2:
                        print(f"    Row 2: {df.iloc[2].tolist()}")
                    if len(df) > 3:
                        print(f"    ... (showing first 3 rows of {len(df)} total rows)")
                    
                    # Check if this looks like header/footer text vs actual data
                    first_row_text = ' '.join(str(cell) for cell in df.iloc[0] if str(cell).strip())
                    print(f"    First row text: {first_row_text[:150]}...")
                    
                    # Look for signs this might be header/footer rather than data
                    header_keywords = ['ALL SHEETS', 'MAINTAINED', 'REVISION', 'DRAWING', 'CONFIDENTIAL', 'EXPORT CONTROL', 'TITLE', 'CABLE ASSEMBLY']
                    is_likely_header = any(keyword in first_row_text.upper() for keyword in header_keywords)
                    
                    if is_likely_header:
                        print(f"    ⚠️  This table appears to contain header/footer text, not data")
                    else:
                        print(f"    ✅ This table appears to contain actual data")
                    
                    # More lenient criteria for ROI extraction - user specifically selected this area
                    # If we have few tables, be very permissive since user selected this specific area
                    total_tables_found = len(roi_tables)
                    if total_tables_found < 5:
                        # Very permissive criteria when few tables found
                        if not df.empty and df.shape[0] >= 1 and df.shape[1] >= 1:
                            print(f"    Table {i+1}: {df.shape[0]}×{df.shape[1]} (accuracy: {accuracy:.1f}%) - ✅ (ROI - permissive)")
                            tables.append(df)
                        else:
                            print(f"    Table {i+1}: {df.shape[0]}×{df.shape[1]} (accuracy: {accuracy:.1f}%) - ❌ (empty)")
                    else:
                        # Normal criteria when many tables found
                        if not df.empty and df.shape[0] >= 1 and df.shape[1] >= 1 and accuracy > 10:
                            print(f"    Table {i+1}: {df.shape[0]}×{df.shape[1]} (accuracy: {accuracy:.1f}%) - ✅ (ROI)")
                            tables.append(df)
                        else:
                            print(f"    Table {i+1}: {df.shape[0]}×{df.shape[1]} (accuracy: {accuracy:.1f}%) - ❌")
            else:
                print("    ❌ ROI lattice found no tables")
                
        except Exception as e:
            print(f"    ❌ ROI lattice failed: {e}")
            
        # If we got header/footer text instead of data, try with expanded area
        if tables and len(tables) == 1:
            # Check if the single table we found looks like header/footer
            df = tables[0]
            first_row_text = ' '.join(str(cell) for cell in df.iloc[0] if str(cell).strip())
            header_keywords = ['ALL SHEETS', 'MAINTAINED', 'REVISION', 'DRAWING', 'CONFIDENTIAL', 'EXPORT CONTROL', 'TITLE', 'CABLE ASSEMBLY']
            is_likely_header = any(keyword in first_row_text.upper() for keyword in header_keywords)
            
            if is_likely_header:
                print("🔄 Detected header/footer text, trying with expanded ROI area...")
                tables.clear()  # Clear the header/footer table
                
                # Try with more expanded area
                expanded_table_areas = []
                for area in roi_areas:
                    if len(area) == 4:
                        # Bigger padding for expanded search
                        expanded_padding = 20  # 20 points padding
                        pdf_left = max(0, area[1] - expanded_padding)
                        pdf_top = max(0, page_height - area[2] - expanded_padding)
                        pdf_right = min(page_width, area[3] + expanded_padding)
                        pdf_bottom = min(page_height, page_height - area[0] + expanded_padding)
                        
                        expanded_area = f"{pdf_left},{pdf_top},{pdf_right},{pdf_bottom}"
                        expanded_table_areas.append(expanded_area)
                        print(f"    Expanded area: {expanded_area}")
                
                try:
                    expanded_tables = camelot.read_pdf(
                        pdf_path,
                        pages=pages_str,
                        flavor='lattice',
                        table_areas=expanded_table_areas,
                        split_text=True,
                        flag_size=True,
                        strip_text='\n'
                    )
                    
                    if expanded_tables and len(expanded_tables) > 0:
                        print(f"    ✅ Expanded ROI found {len(expanded_tables)} tables")
                        
                        for i, camelot_table in enumerate(expanded_tables):
                            df = camelot_table.df
                            accuracy = camelot_table.accuracy
                            
                            # Check if this looks like actual data
                            first_row_text = ' '.join(str(cell) for cell in df.iloc[0] if str(cell).strip())
                            is_likely_header = any(keyword in first_row_text.upper() for keyword in header_keywords)
                            
                            if not is_likely_header and not df.empty and df.shape[0] >= 1 and df.shape[1] >= 1:
                                print(f"    Table {i+1}: {df.shape[0]}×{df.shape[1]} (accuracy: {accuracy:.1f}%) - ✅ (expanded ROI)")
                                tables.append(df)
                            else:
                                print(f"    Table {i+1}: {df.shape[0]}×{df.shape[1]} (accuracy: {accuracy:.1f}%) - ❌ (still header/footer)")
                    else:
                        print("    ❌ Expanded ROI found no tables")
                        
                except Exception as e:
                    print(f"    ❌ Expanded ROI failed: {e}")
        
        # Try stream method with ROI areas if lattice failed
        if not tables:
            try:
                print(f"🔧 Calling camelot.read_pdf (stream) with table_areas={table_areas}")
                
                # Try multiple stream configurations to improve column detection
                stream_configs = [
                    # Config 1: Default stream settings
                    {
                        'edge_tol': 500,
                        'row_tol': 10,
                        'column_tol': 0,
                        'label': 'default'
                    },
                    # Config 2: More sensitive column detection
                    {
                        'edge_tol': 50,
                        'row_tol': 5,
                        'column_tol': 10,
                        'label': 'sensitive'
                    },
                    # Config 3: Very sensitive column detection
                    {
                        'edge_tol': 20,
                        'row_tol': 2,
                        'column_tol': 5,
                        'label': 'very sensitive'
                    }
                ]
                
                for config in stream_configs:
                    print(f"🔧 Trying stream with {config['label']} settings...")
                    
                    roi_stream_tables = camelot.read_pdf(
                        pdf_path,
                        pages=pages_str,
                        flavor='stream',
                        table_areas=table_areas,
                        split_text=True,
                        flag_size=True,
                        strip_text='\n',
                        edge_tol=config['edge_tol'],
                        row_tol=config['row_tol'],
                        column_tol=config['column_tol']
                    )
                    
                    print(f"🔧 Camelot stream ({config['label']}) returned {len(roi_stream_tables) if roi_stream_tables else 0} tables")
                    
                    if roi_stream_tables and len(roi_stream_tables) > 0:
                        print(f"    ✅ ROI stream ({config['label']}) found {len(roi_stream_tables)} tables")
                        
                        for i, camelot_table in enumerate(roi_stream_tables):
                            df = camelot_table.df
                            accuracy = camelot_table.accuracy
                            
                            print(f"    Table {i+1}: {df.shape[0]}×{df.shape[1]} (accuracy: {accuracy:.1f}%)")
                            print(f"    Table {i+1} sample content:")
                            print(df.head(3).to_string())
                            
                            # More lenient criteria for ROI stream extraction
                            total_tables_found = len(roi_stream_tables)
                            if total_tables_found < 5:
                                # Very permissive criteria when few tables found
                                if not df.empty and df.shape[0] >= 1 and df.shape[1] >= 1:
                                    print(f"    Table {i+1}: {df.shape[0]}×{df.shape[1]} (accuracy: {accuracy:.1f}%) - ✅ (ROI stream - permissive)")
                                    tables.append(df)
                                else:
                                    print(f"    Table {i+1}: {df.shape[0]}×{df.shape[1]} (accuracy: {accuracy:.1f}%) - ❌ (empty)")
                            else:
                                # Normal criteria when many tables found
                                if not df.empty and df.shape[0] >= 1 and df.shape[1] >= 1 and accuracy > 5:
                                    print(f"    Table {i+1}: {df.shape[0]}×{df.shape[1]} (accuracy: {accuracy:.1f}%) - ✅ (ROI stream)")
                                    tables.append(df)
                                else:
                                    print(f"    Table {i+1}: {df.shape[0]}×{df.shape[1]} (accuracy: {accuracy:.1f}%) - ❌")
                        
                        # If we found tables with this config, break and use them
                        if tables:
                            print(f"✅ Using tables from {config['label']} stream configuration")
                            break
                    else:
                        print(f"    ❌ ROI stream ({config['label']}) found no tables")
                        
            except Exception as e:
                print(f"    ❌ ROI stream failed: {e}")
        
        # If ROI extraction failed completely, try without ROI as fallback
        if not tables:
            print("🔄 ROI extraction failed, trying without ROI restriction as fallback...")
            try:
                fallback_tables = camelot.read_pdf(
                    pdf_path,
                    pages=pages_str,
                    flavor='stream',
                    split_text=True,
                    flag_size=True,
                    strip_text='\n'
                )
                
                if fallback_tables and len(fallback_tables) > 0:
                    print(f"    ✅ Fallback found {len(fallback_tables)} tables (no ROI)")
                    
                    for i, camelot_table in enumerate(fallback_tables):
                        df = camelot_table.df
                        accuracy = camelot_table.accuracy
                        
                        print(f"    Table {i+1}: {df.shape[0]}×{df.shape[1]} (accuracy: {accuracy:.1f}%)")
                        
                        # Standard criteria for fallback
                        if not df.empty and df.shape[0] >= 2 and df.shape[1] >= 2 and accuracy > 20:
                            print(f"    Table {i+1}: {df.shape[0]}×{df.shape[1]} (accuracy: {accuracy:.1f}%) - ✅ (fallback)")
                            tables.append(df)
                        else:
                            print(f"    Table {i+1}: {df.shape[0]}×{df.shape[1]} (accuracy: {accuracy:.1f}%) - ❌")
                else:
                    print("    ❌ Fallback found no tables")
                    
            except Exception as e:
                print(f"    ❌ Fallback extraction failed: {e}")
        
        return tables
        
    except Exception as e:
        print(f"❌ Camelot ROI extraction failed: {e}")
        return []

def get_camelot_table_info(camelot_table):
    """
    Get detailed information about a camelot table object.
    
    Args:
        camelot_table: Camelot table object
        
    Returns:
        dict: Table information
    """
    try:
        return {
            'shape': camelot_table.df.shape,
            'accuracy': camelot_table.accuracy,
            'whitespace': camelot_table.whitespace,
            'order': camelot_table.order,
            'page': camelot_table.page
        }
    except Exception as e:
        return {'error': str(e)}

def extract_tables_with_camelot_advanced(pdf_path, pages, **kwargs):
    """
    Advanced camelot extraction with custom parameters.
    
    Args:
        pdf_path (str): Path to PDF file
        pages (str): Page range to extract from
        **kwargs: Additional camelot parameters
        
    Returns:
        list: List of extracted DataFrames
    """
    print(f"🐪 Advanced Camelot extraction with custom parameters...")
    
    try:
        import camelot
    except ImportError:
        print("❌ Camelot not available")
        return []
    
    try:
        # Default parameters
        default_params = {
            'flavor': 'lattice',
            'split_text': True,
            'flag_size': True,
            'strip_text': '\n',
            'line_scale': 40,
            'copy_text': ['h', 'v'],
            'shift_text': ['l', 't'],
            'process_background': True
        }
        
        # Update with custom parameters
        params = {**default_params, **kwargs}
        
        print(f"🐪 Using parameters: {params}")
        
        # Convert pages parameter
        if isinstance(pages, list):
            pages_str = ','.join(map(str, pages))
        else:
            pages_str = str(pages)
        
        # Extract tables
        camelot_tables = camelot.read_pdf(pdf_path, pages=pages_str, **params)
        
        tables = []
        if camelot_tables and len(camelot_tables) > 0:
            print(f"    ✅ Advanced extraction found {len(camelot_tables)} tables")
            
            for i, camelot_table in enumerate(camelot_tables):
                df = camelot_table.df
                accuracy = camelot_table.accuracy
                
                if not df.empty and df.shape[0] >= 2 and df.shape[1] >= 2:
                    print(f"    Table {i+1}: {df.shape[0]}×{df.shape[1]} (accuracy: {accuracy:.1f}%)")
                    tables.append(df)
        else:
            print("    ❌ Advanced extraction found no tables")
        
        return tables
        
    except Exception as e:
        print(f"❌ Advanced Camelot extraction failed: {e}")
        return []