import tabula
import pandas as pd
import os
import sys
import re
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from pathlib import Path
from pandastable import Table, TableModel
from validation_utils import validate_extracted_tables, handle_common_errors, generate_output_path
import logging

# Configure matplotlib to use non-interactive backend to avoid GUI issues
try:
    import matplotlib
    matplotlib.use('Agg')
except ImportError:
    pass  # matplotlib not available

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
        total_score = header_score + structure_score + strong_positive_score - min(reject_score, 5)  # Cap reject penalty
        min_threshold = 5  # Lower threshold when we have strong BOM indicators
    else:
        # Standard scoring for tables without strong BOM indicators
        total_score = header_score + structure_score - reject_score
        min_threshold = 2
    
    print(f"    BOM Analysis - Header score: {header_score}, Structure score: {structure_score}, Strong BOM score: {strong_positive_score}, Reject score: {reject_score}, Total: {total_score}, Threshold: {min_threshold}")
    
    # Accept if we meet the threshold
    return total_score >= min_threshold

def extract_tables_with_tabula(pdf_path, pages):
    try:
        print(f"\nExtracting tables from pages: {pages} ...")
        print(f"PDF path: {pdf_path}")
        
        # Handle path encoding issues
        try:
            # Ensure the PDF path is properly encoded
            if isinstance(pdf_path, str):
                pdf_path = pdf_path.encode('utf-8', errors='ignore').decode('utf-8')
            print(f"Normalized PDF path: {pdf_path}")
        except Exception as path_error:
            print(f"Warning: Could not normalize PDF path: {path_error}")
        
        # Check if running in PyInstaller environment
        if getattr(sys, 'frozen', False):
            print("Running in PyInstaller environment")
            print(f"JAVA_TOOL_OPTIONS: {os.environ.get('JAVA_TOOL_OPTIONS', 'Not set')}")
        
        # Set environment variables to handle encoding issues
        original_java_options = os.environ.get('JAVA_TOOL_OPTIONS', '')
        original_lang = os.environ.get('LANG', '')
        original_lc_all = os.environ.get('LC_ALL', '')
        
        # Set environment to handle subprocess encoding issues
        os.environ['PYTHONIOENCODING'] = 'utf-8:ignore'
        os.environ['LANG'] = 'C.UTF-8'
        os.environ['LC_ALL'] = 'C.UTF-8'
        
        tables = []
        
        # Method 1: Try with subprocess encoding handling
        try:
            print("Attempting lattice method with subprocess encoding fix...")
            os.environ['JAVA_TOOL_OPTIONS'] = '-Dfile.encoding=UTF-8 -Duser.language=en -Duser.country=US'
            
            # Import here to avoid issues if jpype is not available
            import subprocess
            import tempfile
            import json
            
            # Create a wrapper that handles encoding issues
            tables = tabula.read_pdf(
                pdf_path,
                pages=pages,
                multiple_tables=True,
                lattice=True,
                java_options="-Dfile.encoding=UTF-8 -Duser.language=en -Duser.country=US -Djava.awt.headless=true",
                pandas_options={'header': None}
            )
            print(f"✅ Lattice method with encoding fix extracted {len(tables)} tables")
        except Exception as e:
            print(f"❌ Lattice method failed: {e}")
            
            # Try with different subprocess approach
            try:
                print("Attempting with subprocess encoding workaround...")
                # Set different environment encoding
                os.environ['PYTHONIOENCODING'] = 'latin1:ignore'
                tables = tabula.read_pdf(
                    pdf_path,
                    pages=pages,
                    multiple_tables=True,
                    lattice=True,
                    java_options="-Dfile.encoding=ISO-8859-1 -Duser.language=en -Duser.country=US -Djava.awt.headless=true",
                    pandas_options={'header': None}
                )
                print(f"✅ Subprocess encoding workaround extracted {len(tables)} tables")
            except Exception as e2:
                print(f"❌ Subprocess encoding workaround failed: {e2}")
                tables = []
        
        # Method 2: Last resort - try with completely different approach
        if not tables:
            try:
                print("Attempting last resort method...")
                
                # Try to use camelot as alternative if available
                try:
                    import camelot
                    print("Trying camelot as alternative...")
                    
                    # First try lattice method with better configuration
                    print("  Trying camelot lattice method...")
                    tables_camelot = camelot.read_pdf(
                        pdf_path, 
                        pages=str(pages), 
                        flavor='lattice',
                        split_text=True,   # Split text in cells
                        flag_size=True,    # Use table size for filtering
                        strip_text='\n'    # Strip newlines
                    )
                    
                    if tables_camelot and len(tables_camelot) > 0:
                        # Filter tables by accuracy and size
                        good_tables = []
                        for i, table in enumerate(tables_camelot):
                            accuracy = table.accuracy
                            rows, cols = table.df.shape
                            print(f"  Camelot table {i+1}: accuracy={accuracy:.1f}%, size={rows}x{cols}")
                            
                            # Only keep tables with reasonable accuracy and size
                            if accuracy > 50 and rows >= 3 and cols >= 3:
                                good_tables.append(table.df)
                                print(f"    ✅ Keeping table {i+1} (good accuracy and size)")
                            else:
                                print(f"    ❌ Skipping table {i+1} (low accuracy or too small)")
                        
                        if good_tables:
                            tables = good_tables
                            print(f"✅ Camelot lattice extracted {len(tables)} good tables")
                        else:
                            print("❌ No good quality tables found with lattice method")
                    
                    # If lattice didn't work well, try stream method
                    if not tables:
                        print("  Trying camelot stream method...")
                        tables_camelot = camelot.read_pdf(
                            pdf_path, 
                            pages=str(pages), 
                            flavor='stream',
                            split_text=True,   # Split text in cells
                            flag_size=True,    # Use table size for filtering
                            strip_text='\n',   # Strip newlines
                            edge_tol=500       # Tolerance for edge detection
                        )
                        
                        if tables_camelot and len(tables_camelot) > 0:
                            # Filter tables by accuracy and size
                            good_tables = []
                            for i, table in enumerate(tables_camelot):
                                accuracy = table.accuracy
                                rows, cols = table.df.shape
                                print(f"  Camelot table {i+1}: accuracy={accuracy:.1f}%, size={rows}x{cols}")
                                
                                # Only keep tables with reasonable accuracy and size
                                if accuracy > 30 and rows >= 3 and cols >= 3:
                                    good_tables.append(table.df)
                                    print(f"    ✅ Keeping table {i+1} (acceptable accuracy and size)")
                                else:
                                    print(f"    ❌ Skipping table {i+1} (low accuracy or too small)")
                            
                            if good_tables:
                                tables = good_tables
                                print(f"✅ Camelot stream extracted {len(tables)} good tables")
                            else:
                                print("❌ No good quality tables found with stream method")
                        
                    if not tables:
                        print("❌ Camelot found no suitable tables")
                        
                except ImportError:
                    print("Camelot not available, trying final tabula attempt...")
                    
                    # Final attempt with minimal options
                    os.environ['PYTHONIOENCODING'] = 'utf-8:ignore'
                    tables = tabula.read_pdf(
                        pdf_path,
                        pages=pages,
                        multiple_tables=True,
                        stream=True,
                        guess=True
                    )
                    print(f"✅ Final minimal attempt extracted {len(tables)} tables")
                    
            except Exception as e:
                print(f"❌ Last resort method failed: {e}")
                tables = []
        
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
            
        # Report extraction results
        if tables:
            print(f"🎉 Successfully extracted {len(tables)} tables using one of the fallback methods")
        else:
            print("❌ All extraction methods failed - no tables extracted")
            return []
        
        # Clean and process the extracted tables
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
                        print(f"Table {i+1}: {table.shape[0]} rows, {table.shape[1]} columns - ✅ Looks like BOM")
                        # Show a sample of the cleaned data for debugging
                        print(f"Sample cleaned data from Table {i+1}:")
                        print(table.head(2).to_string())
                        print("---")
                    else:
                        print(f"Table {i+1}: {table.shape[0]} rows, {table.shape[1]} columns - ❌ Doesn't look like BOM, skipping")
                        # Show sample of rejected table for debugging
                        print(f"Sample rejected data from Table {i+1}:")
                        print(table.head(2).to_string())
                        print("---")
        
        tables = cleaned_tables
        
        # Validate extracted tables
        tables_valid, warning_message = validate_extracted_tables(tables)
        
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
        
        print(f"Successfully extracted {len(tables)} tables.")
        return tables
        
    except Exception as e:
        error_message = f"Error using Tabula: {e}"
        friendly_error = handle_common_errors(str(e))
        print(friendly_error)
        
        # Show error in GUI if available
        try:
            messagebox.showerror("Table Extraction Error", friendly_error)
        except:
            pass  # GUI not available, error already printed
        
        return []

def clean_farrell_columns(df):
    """Clean Farrell-specific table formatting, robustly finding the correct header row and removing all rows above it."""
    print(f"\n🔧 FARRELL DEBUG: Original table shape: {df.shape}")
    print(f"🔧 FARRELL DEBUG: First few rows:\n{df.head(8)}")
    if df.empty:
        print("🔧 FARRELL DEBUG: Empty dataframe passed to clean_farrell_columns")
        return df

    # Define header keywords for scoring
    header_keywords = [
        'QTY', 'PART', 'MFG', 'PAF', 'DESCRIPTION', 'DESCRIP', 'COMMENTS', 'ITEM', 'NUMBER', 'INTERNAL', 'MANUFACTURER', 'MPN'
    ]
    best_score = 0
    best_idx = 0
    for idx in range(min(10, len(df))):  # Scan first 10 rows for best header
        row = df.iloc[idx]
        non_empty_cells = row.dropna().astype(str).str.upper().str.strip()
        score = sum(any(kw in cell for kw in header_keywords) for cell in non_empty_cells)
        print(f"🔧 FARRELL DEBUG: Row {idx} header score: {score} - {non_empty_cells.tolist()}")
        if score > best_score:
            best_score = score
            best_idx = idx
    print(f"🔧 FARRELL DEBUG: Selected header row index: {best_idx} (score: {best_score})")

    # Set the detected header row as columns
    new_columns = df.iloc[best_idx].fillna('').astype(str).str.strip()
    for i, col in enumerate(new_columns):
        if col == '' or col == 'nan':
            new_columns.iloc[i] = f'Column_{i}'
    df.columns = new_columns
    # Remove all rows up to and including the header row
    df = df.iloc[best_idx + 1:].reset_index(drop=True)
    print(f"🔧 FARRELL DEBUG: After header extraction - columns: {df.columns.tolist()}")
    print(f"🔧 FARRELL DEBUG: After header extraction - shape: {df.shape}")

    # Remove any duplicate header rows (sometimes headers repeat as first data row)
    df = df[~df.apply(lambda row: row.astype(str).str.strip().tolist() == df.columns.astype(str).tolist(), axis=1)]
    df = df.reset_index(drop=True)

    # Handle "PART NUMBER" renaming
    for col in df.columns:
        if "PART NUMBER" in str(col).upper():
            df.rename(columns={col: "Internal Part Number"}, inplace=True)
            print(f"🔧 FARRELL DEBUG: Renamed '{col}' to 'Internal Part Number'")
            break

    # Find and split the MFG/PART column
    part_col = None
    for col in df.columns:
        col_str = str(col).upper()
        if ("MFG" in col_str or "MANUF" in col_str) and ("PART" in col_str or "PAF" in col_str):
            part_col = col
            break
    print(f"🔧 FARRELL DEBUG: Found MFG/PART column: {part_col}")
    if part_col and len(df) > 0:
        split_cols = df[part_col].astype(str).str.split("/", n=1, expand=True)
        if split_cols.shape[1] == 2:
            df.insert(0, "Manufacturer", split_cols[0].str.strip())
            df.insert(1, "MPN", split_cols[1].str.strip())
            print("🔧 FARRELL DEBUG: ✅ Successfully split MFG/PART column into Manufacturer and MPN")
            df.drop(columns=[part_col], inplace=True)
        else:
            print("🔧 FARRELL DEBUG: ⚠️ Could not split MFG/PART column - keeping original")
    print(f"🔧 FARRELL DEBUG: Final table shape: {df.shape}")
    print(f"🔧 FARRELL DEBUG: Final columns: {df.columns.tolist()}")
    if len(df) > 0:
        print(f"🔧 FARRELL DEBUG: Sample of final data:\n{df.head(2)}")
    print("🔧 FARRELL DEBUG: ===== END FARRELL PROCESSING =====\n")
    return df

def clean_nel_columns(df):
    """Clean NEL-specific table formatting, handling their schematic BOM structure."""
    print(f"\n🔧 NEL DEBUG: Original table shape: {df.shape}")
    print(f"🔧 NEL DEBUG: First few rows:\n{df.head(8)}")
    if df.empty:
        print("🔧 NEL DEBUG: Empty dataframe passed to clean_nel_columns")
        return df

    # Look for the actual "BILL OF MATERIAL" header
    bill_of_material_row = -1
    for idx in range(min(20, len(df))):
        row = df.iloc[idx]
        row_text = ' '.join(row.fillna('').astype(str).str.upper())
        if 'BILL OF MATERIAL' in row_text:
            bill_of_material_row = idx
            print(f"🔧 NEL DEBUG: Found 'BILL OF MATERIAL' at row {idx}")
            break
    
    # If we found the BOM section, start looking for headers from there
    start_search = max(0, bill_of_material_row)
    
    # Define header keywords for NEL BOMs (common in schematics)
    header_keywords = [
        'ITEM', 'QTY', 'QUANTITY', 'PART', 'NUMBER', 'DESCRIPTION', 'DESC', 'REFERENCE', 'REF', 'DESIGNATOR',
        'VALUE', 'PACKAGE', 'FOOTPRINT', 'MANUFACTURER', 'MFG', 'MPN', 'VENDOR', 'SUPPLIER', 'NOTES'
    ]
    
    best_score = 0
    best_idx = start_search
    for idx in range(start_search, min(start_search + 10, len(df))):  # Look within 10 rows of BOM section
        row = df.iloc[idx]
        non_empty_cells = row.dropna().astype(str).str.upper().str.strip()
        score = sum(any(kw in cell for kw in header_keywords) for cell in non_empty_cells)
        print(f"🔧 NEL DEBUG: Row {idx} header score: {score} - {non_empty_cells.tolist()}")
        if score > best_score:
            best_score = score
            best_idx = idx
    
    print(f"🔧 NEL DEBUG: Selected header row index: {best_idx} (score: {best_score})")

    # If we didn't find good headers, this might not be a BOM table
    if best_score < 2:
        print(f"🔧 NEL DEBUG: Low header score ({best_score}), this might not be a BOM table")
        return df  # Return as-is, let the BOM filter handle it

    # Set the detected header row as columns
    new_columns = df.iloc[best_idx].fillna('').astype(str).str.strip()
    for i, col in enumerate(new_columns):
        if col == '' or col == 'nan':
            new_columns.iloc[i] = f'Column_{i}'
    df.columns = new_columns
    # Remove all rows up to and including the header row
    df = df.iloc[best_idx + 1:].reset_index(drop=True)
    print(f"🔧 NEL DEBUG: After header extraction - columns: {df.columns.tolist()}")
    print(f"🔧 NEL DEBUG: After header extraction - shape: {df.shape}")

    # Remove columns that don't have a proper header (NEL-specific)
    original_columns = df.columns.tolist()
    columns_to_keep = []
    columns_to_remove = []
    
    for col in df.columns:
        col_str = str(col).strip()
        # Keep columns that have meaningful headers (not empty, not generic Column_X, not just whitespace)
        if col_str and col_str != 'nan' and not col_str.startswith('Column_') and col_str.strip() != '':
            columns_to_keep.append(col)
        else:
            columns_to_remove.append(col)
    
    if columns_to_remove:
        print(f"🔧 NEL DEBUG: Removing columns without proper headers: {columns_to_remove}")
        df = df[columns_to_keep]
        print(f"🔧 NEL DEBUG: After removing headerless columns - shape: {df.shape}")
        print(f"🔧 NEL DEBUG: After removing headerless columns - columns: {df.columns.tolist()}")
    else:
        print("🔧 NEL DEBUG: No columns without headers found to remove")

    # Remove any duplicate header rows
    df = df[~df.apply(lambda row: row.astype(str).str.strip().tolist() == df.columns.astype(str).tolist(), axis=1)]
    df = df.reset_index(drop=True)
    
    # Remove rows that look like drawing instructions or notes
    instruction_keywords = [
        'CUT BACK', 'REMOVE', 'SHRINK TUBING', 'DRAWING NUMBER', 'HARNESS', 'PRINTED DRAWING',
        'REFERENCE ONLY', 'DOCUMENT CONTROL', 'LATEST REVISION', 'PROPERTY OF', 'DELIVERED ON',
        'EXPRESS CONDITION', 'NOT TO BE DISCLOSED', 'MARK PER', 'CONTINUITY TEST', 'LOCATE AND ATTACH'
    ]
    
    # Filter out instruction rows
    original_length = len(df)
    for keyword in instruction_keywords:
        # Check each row for instruction keywords
        mask = ~df.apply(lambda row: any(keyword in str(cell).upper() for cell in row), axis=1)
        df = df[mask]
    
    if len(df) < original_length:
        print(f"🔧 NEL DEBUG: Removed {original_length - len(df)} instruction/note rows")
    
    df = df.reset_index(drop=True)

    # Handle common NEL column standardization
    column_mapping = {
        'ITEM': 'Item',
        'ITEM NO': 'Item',
        'QTY': 'Quantity',
        'QUANTITY': 'Quantity',
        'PART NUMBER': 'Part Number',
        'PART': 'Part Number',
        'DESCRIPTION': 'Description',
        'DESC': 'Description',
        'REFERENCE': 'Reference',
        'REF': 'Reference',
        'DESIGNATOR': 'Reference',
        'VALUE': 'Value',
        'PACKAGE': 'Package',
        'FOOTPRINT': 'Footprint',
        'MANUFACTURER': 'Manufacturer',
        'MFG': 'Manufacturer',
        'MPN': 'MPN',
        'MFG P/N': 'MPN',
        'VENDOR': 'Vendor',
        'SUPPLIER': 'Supplier',
        'PROTON P/N': 'Proton P/N',
        'NOTES': 'Notes'
    }
    
    # Apply column mapping
    for old_col, new_col in column_mapping.items():
        for col in df.columns:
            if old_col in str(col).upper():
                df.rename(columns={col: new_col}, inplace=True)
                print(f"🔧 NEL DEBUG: Renamed '{col}' to '{new_col}'")
                break

    # Clean Quantity column - remove any text besides numbers (NEL-specific)
    quantity_cols = [col for col in df.columns if 'quantity' in str(col).lower() or col == 'Quantity']
    if quantity_cols:
        for qty_col in quantity_cols:
            print(f"🔧 NEL DEBUG: Cleaning quantity column '{qty_col}'")
            original_values = df[qty_col].copy()
            
            # Function to extract only numbers from text
            def extract_numbers(value):
                if pd.isna(value) or value == '':
                    return ''
                # Convert to string and extract only digits and decimal points
                value_str = str(value).strip()
                # Use regex to find numbers (including decimals)
                numbers = re.findall(r'\d+\.?\d*', value_str)
                if numbers:
                    # Take the first number found
                    return numbers[0]
                else:
                    return ''
            
            # Apply the cleaning function
            df[qty_col] = df[qty_col].apply(extract_numbers)
            
            # Log changes made
            changes_made = sum(1 for orig, new in zip(original_values, df[qty_col]) 
                             if str(orig).strip() != str(new).strip())
            if changes_made > 0:
                print(f"🔧 NEL DEBUG: Cleaned {changes_made} quantity values in '{qty_col}'")
                # Show a few examples of changes
                sample_changes = [(orig, new) for orig, new in zip(original_values, df[qty_col]) 
                                if str(orig).strip() != str(new).strip()][:3]
                for orig, new in sample_changes:
                    print(f"🔧 NEL DEBUG: '{orig}' -> '{new}'")
            else:
                print(f"🔧 NEL DEBUG: No changes needed for quantity column '{qty_col}'")

    print(f"🔧 NEL DEBUG: Final table shape: {df.shape}")
    print(f"🔧 NEL DEBUG: Final columns: {df.columns.tolist()}")
    if len(df) > 0:
        print(f"🔧 NEL DEBUG: Sample of final data:\n{df.head(2)}")
    print("🔧 NEL DEBUG: ===== END NEL PROCESSING =====\n")
    return df

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

def show_table_selector(tables):
    selected = []

    def on_submit():
        print(f"DEBUG: ===== CHECKBOX VALIDATION STARTED =====")
        print(f"DEBUG: var_list has {len(var_list)} variables")
        print(f"DEBUG: tables has {len(tables)} tables")
        print(f"DEBUG: Root window exists: {root.winfo_exists()}")
        
        selected_count = 0
        for i, var in enumerate(var_list):
            try:
                table_status = "EMPTY" if tables[i].empty else f"{tables[i].shape[0]}x{tables[i].shape[1]}"
                is_selected = var.get()
                print(f"DEBUG: Table {i+1} ({table_status}): checkbox value = {is_selected} (type: {type(is_selected)})")
                if is_selected:
                    selected.append(tables[i])
                    selected_count += 1
            except Exception as e:
                print(f"DEBUG: ERROR reading checkbox {i+1}: {e}")
                print(f"DEBUG: Variable type: {type(var)}")
                print(f"DEBUG: Variable master: {getattr(var, 'master', 'No master')}")
        
        print(f"DEBUG: Final selection count: {selected_count}")
        print(f"DEBUG: Selected list length: {len(selected)}")
        print(f"DEBUG: ===== CHECKBOX VALIDATION COMPLETED =====")
        
        # Validate that at least one table is selected
        if selected_count == 0:
            messagebox.showwarning(
                "No Tables Selected", 
                "Please select at least one table to continue.\n\n"
                "Use the checkboxes to select the tables you want to include in the output."
            )
            return  # Don't close the window
        
        # ADDITIONAL DEBUG: Log detailed info about selected tables
        print(f"DEBUG: ===== SELECTED TABLES SUMMARY =====")
        for i, table in enumerate(selected):
            print(f"DEBUG: Selected table {i+1}: shape={table.shape}, first_row={table.iloc[0].to_dict() if len(table) > 0 else 'EMPTY'}")
        print(f"DEBUG: ===== END SUMMARY =====")
        
        root.destroy()

    root = tk.Tk()
    root.title("Select Tables to Keep")
    root.attributes('-fullscreen', True)  # True fullscreen mode
    root.configure(bg='white')  # Set background to white
    
    # Add escape key to exit fullscreen
    root.bind('<Escape>', lambda e: root.attributes('-fullscreen', False))
    root.bind('<F11>', lambda e: root.attributes('-fullscreen', not root.attributes('-fullscreen')))

    container = tk.Frame(root, bg='white')
    container.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    canvas = tk.Canvas(container, bg='white')  # White background for canvas
    scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
    scrollable_frame = tk.Frame(canvas, bg='white')

    scrollable_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
    )

    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    # Add instructions at the top
    instructions_frame = tk.Frame(scrollable_frame, bg='white')
    instructions_frame.pack(fill=tk.X, padx=10, pady=10)
    
    title_label = tk.Label(instructions_frame, text="📋 Table Selection", font=('Arial', 14, 'bold'), bg='white')
    title_label.pack(anchor="w")
    
    instructions_label = tk.Label(
        instructions_frame, 
        text="Please review the tables below and CHECK the ones you want to include in the final output.\n"
             "⚠️  All tables start UNSELECTED - you must check the boxes for tables you want to keep.",
        font=('Arial', 10),
        foreground='blue',
        bg='white',
        justify=tk.LEFT
    )
    instructions_label.pack(anchor="w", pady=(5, 10))
    
    # Control buttons frame
    control_frame = tk.Frame(scrollable_frame, bg='white')
    control_frame.pack(fill=tk.X, padx=10, pady=5)

    var_list = []
    for i, table in enumerate(tables):
        # Create a frame for each table - use tk.Frame for consistency
        frame = tk.LabelFrame(scrollable_frame, text=f"Table {i+1}", bg='white', font=('Arial', 12, 'bold'))
        frame.pack(fill=tk.X, padx=10, pady=10)

        # Handle empty tables but still add a checkbox variable to maintain index alignment
        if table.empty:
            empty_label = tk.Label(frame, text="Empty table", bg='white', font=('Arial', 10))
            empty_label.pack(padx=5, pady=5)
            # Still add a checkbox variable for empty tables to maintain index alignment
            var = tk.BooleanVar(master=root, value=False)  # Default to unchecked for empty tables, explicit master
            # Create a simple frame for the checkbox
            checkbox_frame = tk.Frame(frame, bg='white')
            checkbox_frame.pack(anchor="w", padx=5, pady=5)
            cb = tk.Checkbutton(
                checkbox_frame, 
                text="Select this table (empty)", 
                variable=var, 
                font=('Arial', 11), 
                state='disabled',
                bg='white'
            )
            cb.pack(anchor="w")
            var_list.append(var)
            continue

        # Create a frame for the text widget with scrollbars
        text_frame = tk.Frame(frame, bg='white')
        text_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Create Text widget with scrollbars for table display
        text_widget = tk.Text(
            text_frame,
            wrap=tk.NONE,  # Disable word wrapping to enable horizontal scrolling
            font=('Consolas', 10),  # Monospace font for better table alignment
            bg='white',
            fg='black',
            height=20,  # Fixed height to prevent overwhelming display
            state=tk.DISABLED  # Read-only
        )
        
        # Add scrollbars for the text widget
        v_text_scrollbar = ttk.Scrollbar(text_frame, orient="vertical", command=text_widget.yview)
        h_text_scrollbar = ttk.Scrollbar(text_frame, orient="horizontal", command=text_widget.xview)
        text_widget.configure(yscrollcommand=v_text_scrollbar.set, xscrollcommand=h_text_scrollbar.set)
        
        # Format table as readable text - display ALL rows
        table_text = format_table_as_text(table)
        
        # Insert the formatted text
        text_widget.config(state=tk.NORMAL)
        text_widget.insert(tk.END, table_text)
        text_widget.config(state=tk.DISABLED)
        
        # Grid layout for text widget and scrollbars
        text_widget.grid(row=0, column=0, sticky='nsew')
        v_text_scrollbar.grid(row=0, column=1, sticky='ns')
        h_text_scrollbar.grid(row=1, column=0, sticky='ew')
        
        # Configure grid weights for resizing
        text_frame.grid_rowconfigure(0, weight=1)
        text_frame.grid_columnconfigure(0, weight=1)
        
        # Add summary info
        info_label = tk.Label(frame, text=f"Rows: {len(table)}, Columns: {len(table.columns)} (showing all {len(table)} rows)", bg='white', font=('Arial', 9))
        info_label.pack(anchor="w", padx=5, pady=2)

        # Checkbox to select/deselect table - Fix parent window association
        var = tk.BooleanVar(master=root, value=False)  # Explicitly set master to root window
        
        # Add a callback to test checkbox functionality
        def checkbox_callback():
            print(f"DEBUG: Checkbox {len(var_list)} clicked! New value: {var.get()}")
        
        # Create a simple frame for the checkbox
        checkbox_frame = tk.Frame(frame, bg='white')  # Ensure consistent background
        checkbox_frame.pack(anchor="w", padx=5, pady=5)
        cb = tk.Checkbutton(
            checkbox_frame, 
            text="Select this table", 
            variable=var, 
            font=('Arial', 11),
            bg='white',  # Match background
            activebackground='white',  # Consistent when clicked
            command=checkbox_callback  # Add callback for testing
        )
        cb.pack(anchor="w")
        var_list.append(var)

    # NOW define the control functions after var_list is populated
    def select_all():
        """Select all tables"""
        print(f"DEBUG: select_all() called - var_list has {len(var_list)} variables")
        for i, var in enumerate(var_list):
            try:
                var.set(True)
                print(f"DEBUG: Set checkbox {i+1} to True")
            except Exception as e:
                print(f"DEBUG: Error setting checkbox {i+1}: {e}")
    
    def clear_all():
        """Clear all table selections"""
        print(f"DEBUG: clear_all() called - var_list has {len(var_list)} variables")
        for i, var in enumerate(var_list):
            try:
                var.set(False)
                print(f"DEBUG: Set checkbox {i+1} to False")
            except Exception as e:
                print(f"DEBUG: Error clearing checkbox {i+1}: {e}")
    
    # Add the buttons now that the functions are properly defined
    select_all_btn = tk.Button(control_frame, text="✓ Select All", command=select_all, bg='lightgreen', font=('Arial', 10))
    select_all_btn.pack(side=tk.LEFT, padx=5)
    
    clear_all_btn = tk.Button(control_frame, text="✗ Clear All", command=clear_all, bg='lightcoral', font=('Arial', 10))
    clear_all_btn.pack(side=tk.LEFT, padx=5)
    
    # Add test button to debug checkbox states
    def test_checkboxes():
        print(f"DEBUG: TEST BUTTON - Checkbox states:")
        for i, var in enumerate(var_list):
            try:
                state = var.get()
                print(f"DEBUG: Checkbox {i+1}: {state}")
            except Exception as e:
                print(f"DEBUG: Error reading checkbox {i+1}: {e}")
    
    test_btn = tk.Button(control_frame, text="🔍 Test Checkboxes", command=test_checkboxes, bg='yellow', font=('Arial', 10))
    test_btn.pack(side=tk.LEFT, padx=5)

    # Continue button
    button_frame = tk.Frame(root, bg='white')
    button_frame.pack(pady=20)
    
    continue_button = tk.Button(
        button_frame, 
        text="🚀 Continue with Selected Tables", 
        command=on_submit,
        bg='lightblue',
        font=('Arial', 12, 'bold'),
        padx=20,
        pady=10
    )
    continue_button.pack()
    
    # Add a label showing current selection count
    selection_label = tk.Label(button_frame, text="No tables selected", foreground='red', bg='white', font=('Arial', 10))
    selection_label.pack(pady=(10, 0))
    
    def update_selection_count():
        """Update the selection count display"""
        try:
            count = sum(var.get() for var in var_list)
            if count == 0:
                selection_label.config(text="⚠️  No tables selected", foreground='red')
            elif count == 1:
                selection_label.config(text="✓ 1 table selected", foreground='green')
            else:
                selection_label.config(text=f"✓ {count} tables selected", foreground='green')
            
            # Debug output every 10th update to avoid spam
            if hasattr(update_selection_count, 'debug_counter'):
                update_selection_count.debug_counter += 1
            else:
                update_selection_count.debug_counter = 1
            
            if update_selection_count.debug_counter % 10 == 0:
                print(f"DEBUG: Selection count update #{update_selection_count.debug_counter}: {count} selected")
            
        except Exception as e:
            print(f"DEBUG: Error in update_selection_count: {e}")
            selection_label.config(text="⚠️  Error reading selections", foreground='red')
        
        # Schedule next update
        root.after(500, update_selection_count)
    
    # Start the selection count updates
    update_selection_count()
    
    root.mainloop()
    return selected

def review_and_edit_dataframe(df):
    """Review and edit dataframe with error handling for GUI issues."""
    print(f"📝 REVIEW DEBUG: Starting review window for dataframe with shape {df.shape}")
    
    try:
        # Create the window with better configuration
        root = tk.Tk()
        root.title("Review and Edit Merged BoM Table")
        root.geometry("1200x800")
        
        # Force the window to be on top and get focus
        root.lift()
        root.attributes('-topmost', True)
        root.after(100, lambda: root.attributes('-topmost', False))
        
        # Center the window
        root.update_idletasks()
        x = (root.winfo_screenwidth() // 2) - (1200 // 2)
        y = (root.winfo_screenheight() // 2) - (800 // 2)
        root.geometry(f"1200x800+{x}+{y}")
        
        print(f"📝 REVIEW DEBUG: Window created and positioned")
        
        # Main frame
        main_frame = tk.Frame(root)
        main_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Title
        title_label = tk.Label(main_frame, text="Review and Edit Merged BoM Table", 
                              font=("Arial", 14, "bold"))
        title_label.pack(pady=(0, 10))
        
        # Instructions
        instructions = tk.Label(main_frame, 
                               text="Review the merged table below. You can edit cell values directly in the table.\nClick 'Continue' when you're satisfied with the data.",
                               font=("Arial", 10))
        instructions.pack(pady=(0, 10))
        
        # Table frame
        table_frame = tk.Frame(main_frame)
        table_frame.pack(fill='both', expand=True)

        # Try to create the pandastable with error handling
        try:
            print(f"📝 REVIEW DEBUG: Creating pandastable...")
            pt = Table(table_frame, dataframe=df, showtoolbar=True, showstatusbar=True)
            pt.show()
            print(f"📝 REVIEW DEBUG: Pandastable created successfully")
            
            def on_continue():
                nonlocal df
                try:
                    print(f"📝 REVIEW DEBUG: User clicked Continue - extracting data...")
                    df = pt.model.df.copy()
                    print(f"📝 REVIEW DEBUG: Data extracted successfully - shape: {df.shape}")
                except Exception as e:
                    print(f"📝 REVIEW DEBUG: Could not get edited data: {e}")
                    print("📝 REVIEW DEBUG: Using original data...")
                root.destroy()

            def on_cancel():
                print(f"📝 REVIEW DEBUG: User clicked Cancel - using original data")
                root.destroy()

            # Button frame
            button_frame = tk.Frame(main_frame)
            button_frame.pack(fill='x', pady=(10, 0))
            
            tk.Button(button_frame, text="Continue with Changes", command=on_continue,
                     bg='lightgreen', font=("Arial", 12), padx=20, pady=5).pack(side='left', padx=(0, 10))
            tk.Button(button_frame, text="Cancel (Use Original)", command=on_cancel,
                     bg='lightcoral', font=("Arial", 12), padx=20, pady=5).pack(side='left')
            
        except Exception as table_error:
            print(f"📝 REVIEW DEBUG: Pandastable failed: {table_error}")
            print("📝 REVIEW DEBUG: Using fallback text display...")
            
            # Fallback to simple text display
            text_widget = tk.Text(table_frame, wrap=tk.NONE, font=("Courier", 10))
            h_scrollbar = tk.Scrollbar(table_frame, orient=tk.HORIZONTAL, command=text_widget.xview)
            v_scrollbar = tk.Scrollbar(table_frame, orient=tk.VERTICAL, command=text_widget.yview)
            text_widget.configure(xscrollcommand=h_scrollbar.set, yscrollcommand=v_scrollbar.set)
            
            # Display table data
            text_widget.insert(tk.END, df.to_string(index=False))
            text_widget.config(state=tk.DISABLED)
            
            # Grid layout
            text_widget.grid(row=0, column=0, sticky='nsew')
            h_scrollbar.grid(row=1, column=0, sticky='ew')
            v_scrollbar.grid(row=0, column=1, sticky='ns')
            table_frame.grid_rowconfigure(0, weight=1)
            table_frame.grid_columnconfigure(0, weight=1)
            
            # Simple continue button for fallback
            def on_simple_continue():
                print(f"📝 REVIEW DEBUG: User continued with read-only review")
                root.destroy()
            
            button_frame = tk.Frame(main_frame)
            button_frame.pack(fill='x', pady=(10, 0))
            tk.Button(button_frame, text="Continue", command=on_simple_continue,
                     bg='lightblue', font=("Arial", 12), padx=20, pady=5).pack()
        
        # Force focus and show window
        print(f"📝 REVIEW DEBUG: Showing window and starting mainloop...")
        root.focus_force()
        root.grab_set()  # Make window modal
        root.mainloop()
        print(f"📝 REVIEW DEBUG: Mainloop completed")
        
    except Exception as e:
        print(f"📝 REVIEW DEBUG: Error in review window: {e}")
        print(f"📝 REVIEW DEBUG: Skipping manual review - using data as-is...")
    
    print(f"📝 REVIEW DEBUG: Review completed - returning dataframe with shape: {df.shape}")
    return df

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
            merged_df = clean_farrell_columns(merged_df)
        elif company == "nel":
            print("📊 MERGE DEBUG: Applying NEL-specific table formatting...")
            merged_df = clean_nel_columns(merged_df)
        else:
            print(f"📊 MERGE DEBUG: No company-specific formatting applied (company='{company}')")
            
            # Try to auto-detect company from the data
            merged_text = ' '.join(merged_df.fillna('').astype(str).values.flatten()).upper()
            if 'NEL HYDROGEN' in merged_text or 'PROTON ENERGY' in merged_text:
                print("📊 MERGE DEBUG: Auto-detected NEL company from data, applying NEL formatting...")
                merged_df = clean_nel_columns(merged_df)
            elif 'FARRELL' in merged_text:
                print("📊 MERGE DEBUG: Auto-detected Farrell company from data, applying Farrell formatting...")
                merged_df = clean_farrell_columns(merged_df)

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
        
        print(f"📊 MERGE DEBUG: Final cleaned table shape: {merged_df.shape}")

        # Check if we should show the review window
        # If called from GUI thread, we'll skip the review here and handle it in the GUI
        skip_review = os.environ.get("BOM_SKIP_REVIEW", "false").lower() == "true"
        
        if not skip_review:
            print("Opening manual review window...")
            merged_df = review_and_edit_dataframe(merged_df)
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

    tables = extract_tables_with_tabula(pdf_path, pages)

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
            "Possible solutions:\n"
            "1. Verify the page range contains actual tables\n"
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