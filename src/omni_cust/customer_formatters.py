"""
Customer-specific table formatting functions for BoMination.

This module contains formatting functions tailored to specific customer requirements.
Each function handles the unique table structure and formatting needs of different customers.
"""

import pandas as pd
import re
import logging


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


def clean_generic_columns(df):
    """
    Generic table formatting that can be applied to any customer's BOM tables.
    This is a fallback when no specific customer formatter is available.
    """
    print(f"\n🔧 GENERIC DEBUG: Original table shape: {df.shape}")
    if df.empty:
        print("🔧 GENERIC DEBUG: Empty dataframe passed to clean_generic_columns")
        return df

    # Define common BOM header keywords
    header_keywords = [
        'ITEM', 'QTY', 'QUANTITY', 'PART', 'NUMBER', 'DESCRIPTION', 'DESC', 'REFERENCE', 'REF',
        'MANUFACTURER', 'MFG', 'MPN', 'VENDOR', 'SUPPLIER', 'NOTES', 'COMMENTS'
    ]
    
    # Find the best header row (similar to other formatters)
    best_score = 0
    best_idx = 0
    for idx in range(min(10, len(df))):
        row = df.iloc[idx]
        non_empty_cells = row.dropna().astype(str).str.upper().str.strip()
        score = sum(any(kw in cell for kw in header_keywords) for cell in non_empty_cells)
        print(f"🔧 GENERIC DEBUG: Row {idx} header score: {score} - {non_empty_cells.tolist()}")
        if score > best_score:
            best_score = score
            best_idx = idx
    
    print(f"🔧 GENERIC DEBUG: Selected header row index: {best_idx} (score: {best_score})")

    # If we found a good header row, use it
    if best_score >= 2:
        new_columns = df.iloc[best_idx].fillna('').astype(str).str.strip()
        for i, col in enumerate(new_columns):
            if col == '' or col == 'nan':
                new_columns.iloc[i] = f'Column_{i}'
        df.columns = new_columns
        df = df.iloc[best_idx + 1:].reset_index(drop=True)
        print(f"🔧 GENERIC DEBUG: After header extraction - columns: {df.columns.tolist()}")
        print(f"🔧 GENERIC DEBUG: After header extraction - shape: {df.shape}")

    # Remove any duplicate header rows
    df = df[~df.apply(lambda row: row.astype(str).str.strip().tolist() == df.columns.astype(str).tolist(), axis=1)]
    df = df.reset_index(drop=True)

    # Basic column standardization
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
        'MANUFACTURER': 'Manufacturer',
        'MFG': 'Manufacturer',
        'MPN': 'MPN',
        'MFG P/N': 'MPN',
        'VENDOR': 'Vendor',
        'SUPPLIER': 'Supplier',
        'NOTES': 'Notes'
    }
    
    # Apply column mapping
    for old_col, new_col in column_mapping.items():
        for col in df.columns:
            if old_col in str(col).upper():
                df.rename(columns={col: new_col}, inplace=True)
                print(f"🔧 GENERIC DEBUG: Renamed '{col}' to '{new_col}'")
                break

    print(f"🔧 GENERIC DEBUG: Final table shape: {df.shape}")
    print(f"🔧 GENERIC DEBUG: Final columns: {df.columns.tolist()}")
    if len(df) > 0:
        print(f"🔧 GENERIC DEBUG: Sample of final data:\n{df.head(2)}")
    print("🔧 GENERIC DEBUG: ===== END GENERIC PROCESSING =====\n")
    return df


def clean_primetals_columns(df):
    """Clean Primetals-specific table formatting, handling dual-column BOM structures."""
    print(f"\n🔧 PRIMETALS DEBUG: Original table shape: {df.shape}")
    print(f"🔧 PRIMETALS DEBUG: Current column names: {df.columns.tolist()}")
    print(f"🔧 PRIMETALS DEBUG: First few rows:\n{df.head(8)}")
    
    if df.empty:
        print("🔧 PRIMETALS DEBUG: Empty dataframe passed to clean_primetals_columns")
        return df

    # Check if the table already has proper BOM headers
    current_columns = [str(col).upper() for col in df.columns]
    bom_headers = ['ITEM', 'MFG', 'DESCRIPTION', 'QTY']
    # Check for either MPN or MFGPART as the part number column
    part_number_header = 'MPN' in current_columns or 'MFGPART' in current_columns
    
    # If we already have proper BOM headers, don't try to extract new ones
    if all(header in current_columns for header in bom_headers) and part_number_header:
        print("🔧 PRIMETALS DEBUG: Table already has proper BOM headers, skipping header extraction")
        
        # Just clean up the data and remove any unwanted rows
        original_length = len(df)
        
        # Remove rows that contain company/confidential information
        confidential_keywords = [
            'PRIMETALS TECHNOLOGIES', 'CONFIDENTIAL', 'PROPRIETARY', 'INTERNAL USE ONLY',
            'NOT FOR DISTRIBUTION', 'COMPANY CONFIDENTIAL'
        ]
        
        # Filter out confidential rows
        for keyword in confidential_keywords:
            # Check each row for confidential keywords
            mask = ~df.apply(lambda row: any(keyword in str(cell).upper() for cell in row), axis=1)
            df = df[mask]
        
        if len(df) < original_length:
            print(f"🔧 PRIMETALS DEBUG: Removed {original_length - len(df)} confidential/header rows")
        
        df = df.reset_index(drop=True)
        
        # Rename MFGPART to MPN for OEMSecrets compatibility
        if 'MFGPART' in df.columns:
            df.rename(columns={'MFGPART': 'MPN'}, inplace=True)
            print("🔧 PRIMETALS DEBUG: Renamed 'MFGPART' to 'MPN' for OEMSecrets compatibility")
        
        # Clean data for OEMSecrets compatibility
        if len(df) > 0:
            print("🔧 PRIMETALS DEBUG: Cleaning data for OEMSecrets compatibility...")
            
            # Clean quantity fields - remove non-numeric characters
            quantity_cols = [col for col in df.columns if any(qty_name in str(col).upper() for qty_name in ['QTY', 'QUANTITY', 'QUAN'])]
            for col in quantity_cols:
                if col in df.columns:
                    original_count = len(df[col].dropna())
                    # Extract only numeric characters and decimal points
                    df[col] = df[col].astype(str).str.replace(r'[^\d.]', '', regex=True)
                    # Remove empty strings and convert to proper format
                    df[col] = df[col].replace('', '1')  # Default to 1 if empty
                    # Handle multiple decimal points - keep only the first one
                    df[col] = df[col].str.replace(r'\.(?=.*\.)', '', regex=True)
                    # Convert to numeric, invalid entries become NaN
                    df[col] = pd.to_numeric(df[col], errors='coerce')
                    # Fill NaN with 1 (default quantity)
                    df[col] = df[col].fillna(1)
                    # Convert to integer if whole number, otherwise keep as float
                    df[col] = df[col].apply(lambda x: int(x) if x == int(x) else x)
                    cleaned_count = len(df[col].dropna())
                    print(f"🔧 PRIMETALS DEBUG: Cleaned {col} column - {original_count} entries processed, {cleaned_count} valid")
            
            # Clean part number fields - remove excessive whitespace and special characters that might cause issues
            part_cols = [col for col in df.columns if any(part_name in str(col).upper() for part_name in ['PART', 'MPN', 'MFGPART'])]
            for col in part_cols:
                if col in df.columns:
                    original_count = len(df[col].dropna())
                    # Remove excessive whitespace
                    df[col] = df[col].astype(str).str.strip()
                    # Replace multiple spaces with single space
                    df[col] = df[col].str.replace(r'\s+', ' ', regex=True)
                    # Remove leading/trailing special characters that might cause issues
                    df[col] = df[col].str.replace(r'^[^\w\d]+|[^\w\d]+$', '', regex=True)
                    # Replace empty strings, 'nan', and 'None' with "N/A" for OEMSecrets compatibility
                    df[col] = df[col].replace(['', 'nan', 'None', 'NaN'], 'N/A')
                    cleaned_count = len(df[col][df[col] != 'N/A'])
                    print(f"🔧 PRIMETALS DEBUG: Cleaned {col} column - {original_count} entries processed, {cleaned_count} valid")
            
            # Clean manufacturer fields - standardize formatting
            mfg_cols = [col for col in df.columns if any(mfg_name in str(col).upper() for mfg_name in ['MFG', 'MANUFACTURER', 'MANUF'])]
            for col in mfg_cols:
                if col in df.columns:
                    original_count = len(df[col].dropna())
                    # Remove excessive whitespace
                    df[col] = df[col].astype(str).str.strip()
                    # Replace multiple spaces with single space
                    df[col] = df[col].str.replace(r'\s+', ' ', regex=True)
                    # Standardize common manufacturer names
                    df[col] = df[col].str.replace(r'(?i)^siemens.*', 'SIEMENS', regex=True)
                    df[col] = df[col].str.replace(r'(?i)^abb.*', 'ABB', regex=True)
                    df[col] = df[col].str.replace(r'(?i)^schneider.*', 'SCHNEIDER', regex=True)
                    df[col] = df[col].str.replace(r'(?i)^eaton.*', 'EATON', regex=True)
                    # Replace empty strings, 'nan', and 'None' with "N/A" for OEMSecrets compatibility
                    df[col] = df[col].replace(['', 'nan', 'None', 'NaN'], 'N/A')
                    cleaned_count = len(df[col][df[col] != 'N/A'])
                    print(f"🔧 PRIMETALS DEBUG: Cleaned {col} column - {original_count} entries processed, {cleaned_count} valid")
            
            # Clean description fields - remove excessive whitespace and standardize
            desc_cols = [col for col in df.columns if any(desc_name in str(col).upper() for desc_name in ['DESC', 'DESCRIPTION'])]
            for col in desc_cols:
                if col in df.columns:
                    original_count = len(df[col].dropna())
                    # Remove excessive whitespace
                    df[col] = df[col].astype(str).str.strip()
                    # Replace multiple spaces with single space
                    df[col] = df[col].str.replace(r'\s+', ' ', regex=True)
                    # Replace empty strings, 'nan', and 'None' with "N/A" for OEMSecrets compatibility
                    df[col] = df[col].replace(['', 'nan', 'None', 'NaN'], 'N/A')
                    cleaned_count = len(df[col][df[col] != 'N/A'])
                    print(f"🔧 PRIMETALS DEBUG: Cleaned {col} column - {original_count} entries processed, {cleaned_count} valid")
            
            # Fill all remaining empty cells with "N/A" to ensure OEMSecrets processes all rows
            print("🔧 PRIMETALS DEBUG: Filling any remaining empty cells with 'N/A' for OEMSecrets compatibility...")
            df = df.fillna('N/A')
            df = df.replace(['', 'nan', 'None', 'NaN'], 'N/A')
            
            # Set QTY to 0 for rows where MFG, MPN, and DESCRIPTION are all "N/A"
            # This prevents OEMSecrets from adding cost for non-existent parts
            if 'MFG' in df.columns and 'MPN' in df.columns and 'DESCRIPTION' in df.columns and 'QTY' in df.columns:
                na_mask = (df['MFG'] == 'N/A') & (df['MPN'] == 'N/A') & (df['DESCRIPTION'] == 'N/A')
                rows_to_zero = na_mask.sum()
                if rows_to_zero > 0:
                    df.loc[na_mask, 'QTY'] = 0
                    print(f"🔧 PRIMETALS DEBUG: Set QTY to 0 for {rows_to_zero} rows where MFG, MPN, and DESCRIPTION are all 'N/A'")
            
            # Don't remove rows with missing critical data - preserve exact PDF structure
            print("🔧 PRIMETALS DEBUG: Preserving all rows to match PDF structure exactly")
            
            df = df.reset_index(drop=True)
            print("🔧 PRIMETALS DEBUG: Data cleaning completed for OEMSecrets compatibility")
        
        print(f"🔧 PRIMETALS DEBUG: Final table shape: {df.shape}")
        print(f"🔧 PRIMETALS DEBUG: Final columns: {df.columns.tolist()}")
        if len(df) > 0:
            print(f"🔧 PRIMETALS DEBUG: Sample of final data:\n{df.head(2)}")
        print("🔧 PRIMETALS DEBUG: ===== END PRIMETALS PROCESSING =====\n")
        
        return df

    # Check if this looks like a dual-column BOM
    if df.shape[1] >= 8:
        # Look for header patterns to identify column groups
        header_row = df.iloc[0] if len(df) > 0 else pd.Series()
        header_str = ' '.join(str(cell) for cell in header_row)
        
        print(f"🔧 PRIMETALS DEBUG: Header row: {header_str}")
        
        # Find repeated patterns indicating dual columns
        if 'ITEM' in header_str and header_str.count('ITEM') >= 2:
            print("🔧 PRIMETALS DEBUG: Detected dual-column BOM table - splitting into individual parts")
            
            # Find the column indices for each side
            left_cols = []
            right_cols = []            # For dual-column BOM, we expect: ITEM, MFG, MPN, DESCRIPTION, QTY on each side
            header_list = [str(cell).strip().upper() for cell in header_row]
            
            # Find all ITEM columns (there should be 2)
            item_positions = [i for i, h in enumerate(header_list) if 'ITEM' in h]
            print(f"🔧 PRIMETALS DEBUG: ITEM positions found: {item_positions}")
            
            if len(item_positions) >= 2:
                # Use the ITEM positions to determine left and right column groups
                left_start = item_positions[0]
                right_start = item_positions[1]
                
                # Standard BOM columns: ITEM, MFG, MPN, DESCRIPTION, QTY
                left_cols = list(range(left_start, min(left_start + 5, len(header_list))))
                right_cols = list(range(right_start, min(right_start + 5, len(header_list))))
                
                # Remove any columns that don't exist
                left_cols = [i for i in left_cols if i < len(header_list)]
                right_cols = [i for i in right_cols if i < len(header_list)]
                
                print(f"🔧 PRIMETALS DEBUG: Left columns: {left_cols}")
                print(f"🔧 PRIMETALS DEBUG: Right columns: {right_cols}")
                
                if len(left_cols) >= 4 and len(right_cols) >= 4:
                    # Extract data from both sides
                    left_data = df.iloc[:, left_cols].copy()
                    right_data = df.iloc[:, right_cols].copy()
                    
                    # Standardize column names
                    standard_cols = ['ITEM', 'MFG', 'MPN', 'DESCRIPTION', 'QTY']
                    left_data.columns = standard_cols[:len(left_data.columns)]
                    right_data.columns = standard_cols[:len(right_data.columns)]
                    
                    # Remove header rows and empty rows
                    left_data = left_data[1:].reset_index(drop=True)  # Skip header
                    right_data = right_data[1:].reset_index(drop=True)  # Skip header
                    
                    # Filter out empty rows
                    left_data = left_data[left_data['ITEM'].notna() & (left_data['ITEM'].astype(str).str.strip() != '')].copy()
                    right_data = right_data[right_data['ITEM'].notna() & (right_data['ITEM'].astype(str).str.strip() != '')].copy()
                    
                    print(f"🔧 PRIMETALS DEBUG: Left side: {len(left_data)} parts")
                    print(f"🔧 PRIMETALS DEBUG: Right side: {len(right_data)} parts")
                    
                    # Combine both sides
                    combined_data = pd.concat([left_data, right_data], ignore_index=True)
                    
                    # Clean up the data
                    combined_data = combined_data.dropna(subset=['ITEM'])
                    combined_data = combined_data[combined_data['ITEM'].astype(str).str.strip() != '']
                    
                    print(f"🔧 PRIMETALS DEBUG: Combined: {len(combined_data)} parts")
                    
                    # Clean data for OEMSecrets compatibility
                    if len(combined_data) > 0:
                        print("🔧 PRIMETALS DEBUG: Cleaning dual-column data for OEMSecrets compatibility...")
                        
                        # Clean quantity fields - remove non-numeric characters
                        quantity_cols = [col for col in combined_data.columns if any(qty_name in str(col).upper() for qty_name in ['QTY', 'QUANTITY', 'QUAN'])]
                        for col in quantity_cols:
                            if col in combined_data.columns:
                                original_count = len(combined_data[col].dropna())
                                # Extract only numeric characters and decimal points
                                combined_data[col] = combined_data[col].astype(str).str.replace(r'[^\d.]', '', regex=True)
                                # Remove empty strings and convert to proper format
                                combined_data[col] = combined_data[col].replace('', '1')  # Default to 1 if empty
                                # Handle multiple decimal points - keep only the first one
                                combined_data[col] = combined_data[col].str.replace(r'\.(?=.*\.)', '', regex=True)
                                # Convert to numeric, invalid entries become NaN
                                combined_data[col] = pd.to_numeric(combined_data[col], errors='coerce')
                                # Fill NaN with 1 (default quantity)
                                combined_data[col] = combined_data[col].fillna(1)
                                # Convert to integer if whole number, otherwise keep as float
                                combined_data[col] = combined_data[col].apply(lambda x: int(x) if x == int(x) else x)
                                cleaned_count = len(combined_data[col].dropna())
                                print(f"🔧 PRIMETALS DEBUG: Cleaned {col} column - {original_count} entries processed, {cleaned_count} valid")
                        
                        # Clean part number fields
                        part_cols = [col for col in combined_data.columns if any(part_name in str(col).upper() for part_name in ['PART', 'MPN', 'MFGPART'])]
                        for col in part_cols:
                            if col in combined_data.columns:
                                original_count = len(combined_data[col].dropna())
                                combined_data[col] = combined_data[col].astype(str).str.strip()
                                combined_data[col] = combined_data[col].str.replace(r'\s+', ' ', regex=True)
                                combined_data[col] = combined_data[col].str.replace(r'^[^\w\d]+|[^\w\d]+$', '', regex=True)
                                combined_data[col] = combined_data[col].replace(['', 'nan', 'None', 'NaN'], 'N/A')
                                cleaned_count = len(combined_data[col][combined_data[col] != 'N/A'])
                                print(f"🔧 PRIMETALS DEBUG: Cleaned {col} column - {original_count} entries processed, {cleaned_count} valid")
                        
                        # Clean manufacturer fields
                        mfg_cols = [col for col in combined_data.columns if any(mfg_name in str(col).upper() for mfg_name in ['MFG', 'MANUFACTURER', 'MANUF'])]
                        for col in mfg_cols:
                            if col in combined_data.columns:
                                original_count = len(combined_data[col].dropna())
                                combined_data[col] = combined_data[col].astype(str).str.strip()
                                combined_data[col] = combined_data[col].str.replace(r'\s+', ' ', regex=True)
                                combined_data[col] = combined_data[col].str.replace(r'(?i)^siemens.*', 'SIEMENS', regex=True)
                                combined_data[col] = combined_data[col].str.replace(r'(?i)^abb.*', 'ABB', regex=True)
                                combined_data[col] = combined_data[col].str.replace(r'(?i)^schneider.*', 'SCHNEIDER', regex=True)
                                combined_data[col] = combined_data[col].str.replace(r'(?i)^eaton.*', 'EATON', regex=True)
                                combined_data[col] = combined_data[col].replace(['', 'nan', 'None', 'NaN'], 'N/A')
                                cleaned_count = len(combined_data[col][combined_data[col] != 'N/A'])
                                print(f"🔧 PRIMETALS DEBUG: Cleaned {col} column - {original_count} entries processed, {cleaned_count} valid")
                        
                        # Clean description fields
                        desc_cols = [col for col in combined_data.columns if any(desc_name in str(col).upper() for desc_name in ['DESC', 'DESCRIPTION'])]
                        for col in desc_cols:
                            if col in combined_data.columns:
                                original_count = len(combined_data[col].dropna())
                                combined_data[col] = combined_data[col].astype(str).str.strip()
                                combined_data[col] = combined_data[col].str.replace(r'\s+', ' ', regex=True)
                                combined_data[col] = combined_data[col].replace(['', 'nan', 'None', 'NaN'], 'N/A')
                                cleaned_count = len(combined_data[col][combined_data[col] != 'N/A'])
                                print(f"🔧 PRIMETALS DEBUG: Cleaned {col} column - {original_count} entries processed, {cleaned_count} valid")
                        
                        # Fill all remaining empty cells with "N/A" to ensure OEMSecrets processes all rows
                        print("🔧 PRIMETALS DEBUG: Filling any remaining empty cells with 'N/A' for OEMSecrets compatibility...")
                        combined_data = combined_data.fillna('N/A')
                        combined_data = combined_data.replace(['', 'nan', 'None', 'NaN'], 'N/A')
                        
                        # Set QTY to 0 for rows where MFG, MPN, and DESCRIPTION are all "N/A"
                        # This prevents OEMSecrets from adding cost for non-existent parts
                        if 'MFG' in combined_data.columns and 'MPN' in combined_data.columns and 'DESCRIPTION' in combined_data.columns and 'QTY' in combined_data.columns:
                            na_mask = (combined_data['MFG'] == 'N/A') & (combined_data['MPN'] == 'N/A') & (combined_data['DESCRIPTION'] == 'N/A')
                            rows_to_zero = na_mask.sum()
                            if rows_to_zero > 0:
                                combined_data.loc[na_mask, 'QTY'] = 0
                                print(f"🔧 PRIMETALS DEBUG: Set QTY to 0 for {rows_to_zero} rows where MFG, MPN, and DESCRIPTION are all 'N/A'")
                        
                        combined_data = combined_data.reset_index(drop=True)
                        print("🔧 PRIMETALS DEBUG: Dual-column data cleaning completed for OEMSecrets compatibility")
                    
                    print(f"🔧 PRIMETALS DEBUG: Final table shape: {combined_data.shape}")
                    print(f"🔧 PRIMETALS DEBUG: Final columns: {combined_data.columns.tolist()}")
                    
                    return combined_data
    
    # If not a dual-column BOM, process as regular table
    print("🔧 PRIMETALS DEBUG: Processing as regular single-column table")
    
    # Define header keywords for scoring
    header_keywords = [
        'ITEM', 'QTY', 'QUANTITY', 'PART', 'NUMBER', 'DESCRIPTION', 'DESC', 'MANUFACTURER', 'MFG', 'MPN', 'MFGPART'
    ]
    
    best_score = 0
    best_idx = 0
    for idx in range(min(10, len(df))):  # Scan first 10 rows for best header
        row = df.iloc[idx]
        non_empty_cells = row.dropna().astype(str).str.upper().str.strip()
        score = sum(any(kw in cell for kw in header_keywords) for cell in non_empty_cells)
        print(f"🔧 PRIMETALS DEBUG: Row {idx} header score: {score} - {non_empty_cells.tolist()}")
        if score > best_score:
            best_score = score
            best_idx = idx
    
    print(f"🔧 PRIMETALS DEBUG: Selected header row index: {best_idx} (score: {best_score})")

    # Set the detected header row as columns
    new_columns = df.iloc[best_idx].fillna('').astype(str).str.strip()
    for i, col in enumerate(new_columns):
        if col == '' or col == 'nan':
            new_columns.iloc[i] = f'Column_{i}'
    
    df.columns = new_columns
    # Remove all rows up to and including the header row
    df = df.iloc[best_idx + 1:].reset_index(drop=True)
    
    print(f"🔧 PRIMETALS DEBUG: After header extraction - columns: {df.columns.tolist()}")
    print(f"🔧 PRIMETALS DEBUG: After header extraction - shape: {df.shape}")

    # Remove any duplicate header rows
    df = df[~df.apply(lambda row: row.astype(str).str.strip().tolist() == df.columns.astype(str).tolist(), axis=1)]
    df = df.reset_index(drop=True)
    
    # Remove rows that contain company/confidential information
    confidential_keywords = [
        'PRIMETALS TECHNOLOGIES', 'CONFIDENTIAL', 'PROPRIETARY', 'INTERNAL USE ONLY',
        'NOT FOR DISTRIBUTION', 'COMPANY CONFIDENTIAL'
    ]
    
    # Filter out confidential rows
    original_length = len(df)
    for keyword in confidential_keywords:
        # Check each row for confidential keywords
        mask = ~df.apply(lambda row: any(keyword in str(cell).upper() for cell in row), axis=1)
        df = df[mask]
    
    if len(df) < original_length:
        print(f"🔧 PRIMETALS DEBUG: Removed {original_length - len(df)} confidential/header rows")
    
    df = df.reset_index(drop=True)
    
    # Handle common Primetals column standardization
    column_mapping = {
        'ITEM': 'ITEM',
        'ITEM NO': 'ITEM', 
        'QTY': 'QTY',
        'QUANTITY': 'QTY',
        'PART NUMBER': 'MPN',
        'PART': 'MPN',
        'MFGPART': 'MPN',
        'MFG PART': 'MPN',
        'DESCRIPTION': 'DESCRIPTION',
        'DESC': 'DESCRIPTION',
        'MANUFACTURER': 'MFG',
        'MFG': 'MFG',
        'MPN': 'MPN',
        'VENDOR': 'VENDOR',
        'SUPPLIER': 'SUPPLIER'
    }
    
    # Apply column mapping
    for old_col, new_col in column_mapping.items():
        matching_cols = [col for col in df.columns if old_col in str(col).upper()]
        if matching_cols:
            df.rename(columns={matching_cols[0]: new_col}, inplace=True)
            print(f"🔧 PRIMETALS DEBUG: Renamed '{matching_cols[0]}' to '{new_col}'")
    
    # Clean data for OEMSecrets compatibility
    if len(df) > 0:
        print("🔧 PRIMETALS DEBUG: Cleaning data for OEMSecrets compatibility...")
        
        # Clean quantity fields - remove non-numeric characters
        quantity_cols = [col for col in df.columns if any(qty_name in str(col).upper() for qty_name in ['QTY', 'QUANTITY', 'QUAN'])]
        for col in quantity_cols:
            if col in df.columns:
                original_count = len(df[col].dropna())
                # Extract only numeric characters and decimal points
                df[col] = df[col].astype(str).str.replace(r'[^\d.]', '', regex=True)
                # Remove empty strings and convert to proper format
                df[col] = df[col].replace('', '1')  # Default to 1 if empty
                # Handle multiple decimal points - keep only the first one
                df[col] = df[col].str.replace(r'\.(?=.*\.)', '', regex=True)
                # Convert to numeric, invalid entries become NaN
                df[col] = pd.to_numeric(df[col], errors='coerce')
                # Fill NaN with 1 (default quantity)
                df[col] = df[col].fillna(1)
                # Convert to integer if whole number, otherwise keep as float
                df[col] = df[col].apply(lambda x: int(x) if x == int(x) else x)
                cleaned_count = len(df[col].dropna())
                print(f"🔧 PRIMETALS DEBUG: Cleaned {col} column - {original_count} entries processed, {cleaned_count} valid")
        
        # Clean part number fields - remove excessive whitespace and special characters that might cause issues
        part_cols = [col for col in df.columns if any(part_name in str(col).upper() for part_name in ['PART', 'MPN', 'MFGPART'])]
        for col in part_cols:
            if col in df.columns:
                original_count = len(df[col].dropna())
                # Remove excessive whitespace
                df[col] = df[col].astype(str).str.strip()
                # Replace multiple spaces with single space
                df[col] = df[col].str.replace(r'\s+', ' ', regex=True)
                # Remove leading/trailing special characters that might cause issues
                df[col] = df[col].str.replace(r'^[^\w\d]+|[^\w\d]+$', '', regex=True)
                # Replace empty strings, 'nan', and 'None' with "N/A" for OEMSecrets compatibility
                df[col] = df[col].replace(['', 'nan', 'None', 'NaN'], 'N/A')
                cleaned_count = len(df[col][df[col] != 'N/A'])
                print(f"🔧 PRIMETALS DEBUG: Cleaned {col} column - {original_count} entries processed, {cleaned_count} valid")
        
        # Clean manufacturer fields - standardize formatting
        mfg_cols = [col for col in df.columns if any(mfg_name in str(col).upper() for mfg_name in ['MFG', 'MANUFACTURER', 'MANUF'])]
        for col in mfg_cols:
            if col in df.columns:
                original_count = len(df[col].dropna())
                # Remove excessive whitespace
                df[col] = df[col].astype(str).str.strip()
                # Replace multiple spaces with single space
                df[col] = df[col].str.replace(r'\s+', ' ', regex=True)
                # Standardize common manufacturer names
                df[col] = df[col].str.replace(r'(?i)^siemens.*', 'SIEMENS', regex=True)
                df[col] = df[col].str.replace(r'(?i)^abb.*', 'ABB', regex=True)
                df[col] = df[col].str.replace(r'(?i)^schneider.*', 'SCHNEIDER', regex=True)
                df[col] = df[col].str.replace(r'(?i)^eaton.*', 'EATON', regex=True)
                # Replace empty strings, 'nan', and 'None' with "N/A" for OEMSecrets compatibility
                df[col] = df[col].replace(['', 'nan', 'None', 'NaN'], 'N/A')
                cleaned_count = len(df[col][df[col] != 'N/A'])
                print(f"🔧 PRIMETALS DEBUG: Cleaned {col} column - {original_count} entries processed, {cleaned_count} valid")
        
        # Clean description fields - remove excessive whitespace and standardize
        desc_cols = [col for col in df.columns if any(desc_name in str(col).upper() for desc_name in ['DESC', 'DESCRIPTION'])]
        for col in desc_cols:
            if col in df.columns:
                original_count = len(df[col].dropna())
                # Remove excessive whitespace
                df[col] = df[col].astype(str).str.strip()
                # Replace multiple spaces with single space
                df[col] = df[col].str.replace(r'\s+', ' ', regex=True)
                # Replace empty strings, 'nan', and 'None' with "N/A" for OEMSecrets compatibility
                df[col] = df[col].replace(['', 'nan', 'None', 'NaN'], 'N/A')
                cleaned_count = len(df[col][df[col] != 'N/A'])
                print(f"🔧 PRIMETALS DEBUG: Cleaned {col} column - {original_count} entries processed, {cleaned_count} valid")
        
        # Fill all remaining empty cells with "N/A" to ensure OEMSecrets processes all rows
        print("🔧 PRIMETALS DEBUG: Filling any remaining empty cells with 'N/A' for OEMSecrets compatibility...")
        df = df.fillna('N/A')
        df = df.replace(['', 'nan', 'None', 'NaN'], 'N/A')
        
        # Set QTY to 0 for rows where MFG, MPN, and DESCRIPTION are all "N/A"
        # This prevents OEMSecrets from adding cost for non-existent parts
        if 'MFG' in df.columns and 'MPN' in df.columns and 'DESCRIPTION' in df.columns and 'QTY' in df.columns:
            na_mask = (df['MFG'] == 'N/A') & (df['MPN'] == 'N/A') & (df['DESCRIPTION'] == 'N/A')
            rows_to_zero = na_mask.sum()
            if rows_to_zero > 0:
                df.loc[na_mask, 'QTY'] = 0
                print(f"🔧 PRIMETALS DEBUG: Set QTY to 0 for {rows_to_zero} rows where MFG, MPN, and DESCRIPTION are all 'N/A'")
        
        # Don't remove rows with missing critical data - preserve exact PDF structure
        print("🔧 PRIMETALS DEBUG: Preserving all rows to match PDF structure exactly")
        
        df = df.reset_index(drop=True)
        print("🔧 PRIMETALS DEBUG: Data cleaning completed for OEMSecrets compatibility")
    
    print(f"🔧 PRIMETALS DEBUG: Final table shape: {df.shape}")
    print(f"🔧 PRIMETALS DEBUG: Final columns: {df.columns.tolist()}")
    if len(df) > 0:
        print(f"🔧 PRIMETALS DEBUG: Sample of final data:\n{df.head(2)}")
    print("🔧 PRIMETALS DEBUG: ===== END PRIMETALS PROCESSING =====\n")
    
    return df


# Customer formatter registry
CUSTOMER_FORMATTERS = {
    'farrell': clean_farrell_columns,
    'nel': clean_nel_columns,
    'generic': clean_generic_columns,
    'primetals': clean_primetals_columns
}


def apply_customer_formatter(df, customer_name=None):
    """
    Apply customer-specific formatting to a dataframe.
    
    Args:
        df: pandas DataFrame to format
        customer_name: string name of the customer (e.g., 'farrell', 'nel')
                      If None or unknown, uses generic formatting
    
    Returns:
        Formatted pandas DataFrame
    """
    if df.empty:
        return df
    
    # Normalize customer name
    if customer_name:
        customer_name = customer_name.lower().strip()
    
    # Get the appropriate formatter
    formatter = CUSTOMER_FORMATTERS.get(customer_name, clean_generic_columns)
    
    print(f"🔧 CUSTOMER FORMATTER: Applying {customer_name or 'generic'} formatting")
    
    try:
        formatted_df = formatter(df)
        print(f"🔧 CUSTOMER FORMATTER: Successfully applied {customer_name or 'generic'} formatting")
        return formatted_df
    except Exception as e:
        print(f"🔧 CUSTOMER FORMATTER: Error applying {customer_name or 'generic'} formatting: {e}")
        print(f"🔧 CUSTOMER FORMATTER: Falling back to generic formatting")
        return clean_generic_columns(df)


def get_available_customers():
    """
    Get a list of available customer formatters.
    
    Returns:
        List of customer names that have specific formatters
    """
    return [name for name in CUSTOMER_FORMATTERS.keys() if name != 'generic']


def add_customer_formatter(customer_name, formatter_function):
    """
    Add a new customer formatter to the registry.
    
    Args:
        customer_name: string name of the customer
        formatter_function: function that takes a DataFrame and returns a formatted DataFrame
    """
    customer_name = customer_name.lower().strip()
    CUSTOMER_FORMATTERS[customer_name] = formatter_function
    print(f"🔧 CUSTOMER FORMATTER: Added formatter for '{customer_name}'")
