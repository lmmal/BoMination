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
        
        print(f"📊 MERGE DEBUG: Final cleaned table shape: {merged_df.shape}")

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