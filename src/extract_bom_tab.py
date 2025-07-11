import tabula
import pandas as pd
import os
import sys
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from pathlib import Path
from pandastable import Table, TableModel
from validation_utils import validate_extracted_tables, handle_common_errors, generate_output_path
import logging

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
        os.environ['JAVA_TOOL_OPTIONS'] = '-Dfile.encoding=UTF-8 -Duser.language=en -Duser.country=US'
        
        tables = []
        
        # Method 1: Try lattice method first
        try:
            print("Attempting lattice method...")
            tables = tabula.read_pdf(
                pdf_path,
                pages=pages,
                multiple_tables=True,
                lattice=True,
                java_options="-Dfile.encoding=UTF-8 -Duser.language=en -Duser.country=US -Djava.awt.headless=true",
                pandas_options={'header': None}
            )
            print(f"✅ Lattice method extracted {len(tables)} tables")
        except (UnicodeDecodeError, UnicodeError) as unicode_error:
            print(f"❌ Lattice method failed with Unicode error: {unicode_error}")
            print("🔄 Skipping to encoding-specific methods...")
            tables = []
        except Exception as e:
            print(f"❌ Lattice method failed: {e}")
            # Check if it's a UTF-8 related error
            if 'utf-8' in str(e).lower() or 'codec' in str(e).lower():
                print("🔄 Detected encoding issue, skipping to encoding-specific methods...")
                tables = []
            else:
                tables = []
            
        # Method 2: If UTF-8 error detected, try ISO-8859-1 immediately
        if not tables:
            # Check if previous error was UTF-8 related
            try:
                print("Attempting with ISO-8859-1 encoding (Windows-1252 compatible)...")
                os.environ['JAVA_TOOL_OPTIONS'] = '-Dfile.encoding=ISO-8859-1 -Duser.language=en -Duser.country=US'
                tables = tabula.read_pdf(
                    pdf_path,
                    pages=pages,
                    multiple_tables=True,
                    stream=True,
                    guess=True,
                    java_options="-Dfile.encoding=ISO-8859-1 -Duser.language=en -Duser.country=US -Djava.awt.headless=true",
                    pandas_options={'header': None}
                )
                print(f"✅ ISO-8859-1 encoding extracted {len(tables)} tables")
            except Exception as e:
                print(f"❌ ISO-8859-1 encoding failed: {e}")
                tables = []
                
        # Method 3: Try CP1252 encoding (Windows default)
        if not tables:
            try:
                print("Attempting with CP1252 encoding...")
                os.environ['JAVA_TOOL_OPTIONS'] = '-Dfile.encoding=CP1252 -Duser.language=en -Duser.country=US'
                tables = tabula.read_pdf(
                    pdf_path,
                    pages=pages,
                    multiple_tables=True,
                    stream=True,
                    guess=True,
                    java_options="-Dfile.encoding=CP1252 -Duser.language=en -Duser.country=US -Djava.awt.headless=true",
                    pandas_options={'header': None}
                )
                print(f"✅ CP1252 encoding extracted {len(tables)} tables")
            except Exception as e:
                print(f"❌ CP1252 encoding failed: {e}")
                tables = []
        
        # Method 4: Try stream method if encoding methods failed
        if not tables:
            try:
                print("Attempting stream method...")
                os.environ['JAVA_TOOL_OPTIONS'] = '-Dfile.encoding=UTF-8 -Duser.language=en -Duser.country=US'
                tables = tabula.read_pdf(
                    pdf_path,
                    pages=pages,
                    multiple_tables=True,
                    stream=True,
                    guess=False,
                    java_options="-Dfile.encoding=UTF-8 -Duser.language=en -Duser.country=US -Djava.awt.headless=true",
                    pandas_options={'header': None}
                )
                print(f"✅ Stream method extracted {len(tables)} tables")
            except Exception as e:
                print(f"❌ Stream method failed: {e}")
                tables = []
        
        # Method 5: Try stream with auto-detection
        if not tables:
            try:
                print("Attempting stream method with auto-detection...")
                tables = tabula.read_pdf(
                    pdf_path,
                    pages=pages,
                    multiple_tables=True,
                    stream=True,
                    guess=True,
                    java_options="-Dfile.encoding=UTF-8 -Duser.language=en -Duser.country=US -Djava.awt.headless=true",
                    pandas_options={'header': None}
                )
                print(f"✅ Stream with auto-detection extracted {len(tables)} tables")
            except Exception as e:
                print(f"❌ Stream with auto-detection failed: {e}")
                tables = []
                
        # Method 6: Last resort - try without any encoding specification
        if not tables:
            try:
                print("Attempting without encoding specification...")
                os.environ['JAVA_TOOL_OPTIONS'] = '-Duser.language=en -Duser.country=US'
                tables = tabula.read_pdf(
                    pdf_path,
                    pages=pages,
                    multiple_tables=True,
                    stream=True,
                    guess=True,
                    java_options="-Duser.language=en -Duser.country=US -Djava.awt.headless=true"
                    # No pandas_options at all
                )
                print(f"✅ No encoding specification extracted {len(tables)} tables")
            except Exception as e:
                print(f"❌ No encoding specification failed: {e}")
                tables = []
        
        # Restore original JAVA_TOOL_OPTIONS
        if original_java_options:
            os.environ['JAVA_TOOL_OPTIONS'] = original_java_options
        elif 'JAVA_TOOL_OPTIONS' in os.environ:
            del os.environ['JAVA_TOOL_OPTIONS']
            
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
                
                if not table.empty:
                    cleaned_tables.append(table)
                    print(f"Table {i+1}: {table.shape[0]} rows, {table.shape[1]} columns")
                    # Show a sample of the cleaned data for debugging
                    print(f"Sample cleaned data from Table {i+1}:")
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

    # Define header keywords for NEL BOMs (common in schematics)
    header_keywords = [
        'ITEM', 'QTY', 'QUANTITY', 'PART', 'NUMBER', 'DESCRIPTION', 'DESC', 'REFERENCE', 'REF', 'DESIGNATOR',
        'VALUE', 'PACKAGE', 'FOOTPRINT', 'MANUFACTURER', 'MFG', 'MPN', 'VENDOR', 'SUPPLIER'
    ]
    
    best_score = 0
    best_idx = 0
    for idx in range(min(15, len(df))):  # Scan first 15 rows for best header (schematics can have more header info)
        row = df.iloc[idx]
        non_empty_cells = row.dropna().astype(str).str.upper().str.strip()
        score = sum(any(kw in cell for kw in header_keywords) for cell in non_empty_cells)
        print(f"🔧 NEL DEBUG: Row {idx} header score: {score} - {non_empty_cells.tolist()}")
        if score > best_score:
            best_score = score
            best_idx = idx
    print(f"🔧 NEL DEBUG: Selected header row index: {best_idx} (score: {best_score})")

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

    # Remove any duplicate header rows
    df = df[~df.apply(lambda row: row.astype(str).str.strip().tolist() == df.columns.astype(str).tolist(), axis=1)]
    df = df.reset_index(drop=True)

    # Handle common NEL column standardization
    column_mapping = {
        'ITEM': 'Item',
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
        'VENDOR': 'Vendor',
        'SUPPLIER': 'Supplier'
    }
    
    # Apply column mapping
    for old_col, new_col in column_mapping.items():
        for col in df.columns:
            if old_col in str(col).upper():
                df.rename(columns={col: new_col}, inplace=True)
                print(f"🔧 NEL DEBUG: Renamed '{col}' to '{new_col}'")
                break

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
    try:
        root = tk.Tk()
        root.title("Review and Edit Merged BoM Table")
        root.withdraw()  # Hide initially to prevent flashing
        
        frame = tk.Frame(root)
        frame.pack(fill='both', expand=True)

        # Try to create the pandastable with error handling
        try:
            pt = Table(frame, dataframe=df, showtoolbar=True, showstatusbar=True)
            pt.show()
            
            def on_continue():
                nonlocal df
                try:
                    df = pt.model.df.copy()
                except Exception as e:
                    print(f"Warning: Could not get edited data: {e}")
                    print("Using original data...")
                root.destroy()

            tk.Button(root, text="Continue", command=on_continue).pack(pady=10)
            
            # Show the window after everything is set up
            root.deiconify()
            root.mainloop()
            
        except Exception as table_error:
            print(f"Error creating table editor: {table_error}")
            print("Skipping manual review - using data as-is...")
            root.destroy()
            
    except Exception as e:
        print(f"Error in review window: {e}")
        print("Skipping manual review - using data as-is...")
    
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

        print("Opening manual review window...")
        merged_df = review_and_edit_dataframe(merged_df)

        merged_df.to_excel(output_path, index=False, sheet_name=sheet_name)
        print(f"✅ Final cleaned and reviewed table saved to: {output_path}")

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
            messagebox.showerror("No Tables Extracted", error_msg)
        except:
            pass  # GUI not available, error already printed
        
        # Exit with error code to indicate failure
        sys.exit(1)

if __name__ == "__main__":
    main()