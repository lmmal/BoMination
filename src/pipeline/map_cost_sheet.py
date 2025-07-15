import pandas as pd
import os
import sys
from openpyxl import load_workbook
from pathlib import Path

# Support both script and PyInstaller .exe paths
if getattr(sys, 'frozen', False):
    # Running as PyInstaller executable
    SCRIPT_DIR = Path(sys._MEIPASS)
    OMNI_TEMPLATE_PATH = SCRIPT_DIR / "Files" / "OCTF-1539-COST SHEET.xlsx"
else:
    # Running as script - Files folder is in root directory, go up two levels from src/pipeline/
    SCRIPT_DIR = Path(__file__).parent
    OMNI_TEMPLATE_PATH = SCRIPT_DIR.parent.parent / "Files" / "OCTF-1539-COST SHEET.xlsx"

def generate_output_path(input_file_path, suffix, output_directory=None, extension=".xlsx"):
    """Generate output file path based on input file and optional output directory."""
    input_path = Path(input_file_path)
    base_name = input_path.stem
    output_filename = f"{base_name}{suffix}{extension}"
    
    if output_directory and output_directory.strip():
        output_dir = Path(output_directory)
        output_dir.mkdir(parents=True, exist_ok=True)
        output_path = output_dir / output_filename
    else:
        output_path = input_path.parent / output_filename
    
    return str(output_path)

def get_oem_file():
    while True:
        path = input("Enter path to OEMSecrets Excel file (with prices): ").strip()
        if os.path.isfile(path) and path.endswith(".xlsx"):
            return path
        print("Invalid file path.")

def get_merged_file():
    while True:
        path = input("Enter path to merged Excel file (with DESCRIPTION column): ").strip()
        if os.path.isfile(path) and path.endswith(".xlsx"):
            return path
        print("Invalid merged file path.")

def map_and_insert_data(oem_path, merged_path, template_path=OMNI_TEMPLATE_PATH):
    df_oem = pd.read_excel(oem_path, keep_default_na=False, na_values=[''])
    df_merged = pd.read_excel(merged_path, keep_default_na=False, na_values=[''])

    # Get the user's selected company from environment variable
    user_company = os.environ.get("BOM_COMPANY", "").lower()
    print(f"📊 User selected company: {user_company}")

    # Pull description column if present
    if "DESCRIPTION" in df_merged.columns:
        df_oem["DESCRIPTION"] = df_merged["DESCRIPTION"]
    elif "Description" in df_merged.columns:  # NEL format
        df_oem["DESCRIPTION"] = df_merged["Description"]
    else:
        print("⚠️ DESCRIPTION column not found in merged file. Skipping description.")
        df_oem["DESCRIPTION"] = ""

    # Pull PROTON P/N column from merged file for customer part number (NEL only)
    if user_company == "nel":
        if "Proton P/N" in df_merged.columns:
            df_oem["PROTON P/N"] = df_merged["Proton P/N"]
            print("✅ Added PROTON P/N column from merged file (NEL format)")
        else:
            print("⚠️ PROTON P/N column not found in merged file. Skipping customer part number.")
            df_oem["PROTON P/N"] = ""
    else:
        print("📊 Skipping PROTON P/N mapping (not NEL format)")
        df_oem["PROTON P/N"] = ""

    # Pull additional columns from merged file for Primetals format
    if user_company == "primetals":
        # Pull MPN (part number) from merged file
        if "MPN" in df_merged.columns:
            df_oem["MPN"] = df_merged["MPN"]
            print("✅ Added MPN column from merged file (Primetals format)")
        else:
            print("⚠️ MPN column not found in merged file. Skipping part number.")
            df_oem["MPN"] = ""
        
        # Pull QTY (quantity) from merged file
        if "QTY" in df_merged.columns:
            df_oem["QTY"] = df_merged["QTY"]
            print("✅ Added QTY column from merged file (Primetals format)")
        else:
            print("⚠️ QTY column not found in merged file. Skipping quantity.")
            df_oem["QTY"] = ""
        
        # Pull MFG (manufacturer) from merged file
        if "MFG" in df_merged.columns:
            df_oem["MFG"] = df_merged["MFG"]
            print("✅ Added MFG column from merged file (Primetals format)")
        else:
            print("⚠️ MFG column not found in merged file. Skipping manufacturer.")
            df_oem["MFG"] = ""
        
        # Pull ITEM from merged file
        if "ITEM" in df_merged.columns:
            df_oem["ITEM"] = df_merged["ITEM"]
            print("✅ Added ITEM column from merged file (Primetals format)")
        else:
            print("⚠️ ITEM column not found in merged file. Skipping item number.")
            df_oem["ITEM"] = ""

    # Detect company format based on column names
    oem_columns = df_oem.columns.tolist()
    merged_columns = df_merged.columns.tolist()
    
    print(f"📊 OEM file columns: {oem_columns}")
    print(f"📊 Merged file columns: {merged_columns}")
    
    # Use user's selected company for mapping instead of auto-detection
    if user_company == "nel":
        print("📊 Using NEL format - based on user selection")
        column_mapping = {
            "Internal Reference": "OMNI PART #",
            "Part Number": "COMMERCIAL PART#",
            "Quantity for Single BOM": "UNIT QTY",
            "Extended Quantity for 1 BOM": "EXT QTY",
            "Manufacturer": "MFR",
            "Distributor": "SUPPLIER / NOTES",
            "Minimum Order": "MIN ORDER",
            "Unit Price in USD": "COST EACH",
            "Lead Time on Additional Stock in Weeks": "LEAD TIME (WEEKS)",
            "PROTON P/N": "CUST PART #",  # Pull from merged file for NEL
            "DESCRIPTION": "DESCRIPTION"
        }
    elif user_company == "primetals":
        print("📊 Using Primetals format - based on user selection")
        column_mapping = {
            "ITEM": "ITEM",
            "MFG": "MFR",
            "MPN": "COMMERCIAL PART#",
            "QTY": "UNIT QTY",
            "Unit Price in USD": "COST EACH",
            "Lead Time on Additional Stock in Weeks": "LEAD TIME (WEEKS)",
            "Notes": "SUPPLIER / NOTES",
            "DESCRIPTION": "DESCRIPTION"
        }
    else:
        print("📊 Using Farrell format - based on user selection")
        # Map the actual column names from the Farrell file
        column_mapping = {
            "Item": "ITEM",
            "Manufacturer": "MFR",
            "MPN": "COMMERCIAL PART#",
            "Quantity": "UNIT QTY",
            "Unit Price in USD": "COST EACH",
            "Lead Time on Additional Stock in Weeks": "LEAD TIME (WEEKS)",
            "Notes": "SUPPLIER / NOTES",
            "DESCRIPTION": "DESCRIPTION"
        }

    print(f"📊 Using column mapping: {column_mapping}")

    df_renamed = df_oem.rename(columns={k: v for k, v in column_mapping.items() if k in df_oem.columns})
    mapped_columns = list(column_mapping.values())
    df_out = df_renamed[[col for col in mapped_columns if col in df_renamed.columns]].copy()

    # Fill any remaining NaN values with "N/A" to ensure consistency
    df_out = df_out.fillna("N/A")
    df_out = df_out.replace("", "N/A")

    # Add ITEM numbers (sequential numbering)
    if "ITEM" not in df_out.columns:
        df_out["ITEM"] = range(1, len(df_out) + 1)

    print(f"📊 Final data shape: {df_out.shape}")
    print(f"📊 Final columns: {df_out.columns.tolist()}")
    if len(df_out) > 0:
        print(f"📊 Sample data:\n{df_out.head(2)}")

    # Validate template file exists before trying to load it
    if not template_path.exists():
        raise FileNotFoundError(
            f"Cost sheet template not found: {template_path}\n\n"
            f"Please ensure the file structure is correct:\n"
            f"BoMination/\n"
            f"  Files/\n"
            f"    OCTF-1539-COST SHEET.xlsx\n"
            f"  src/\n"
            f"    (script files)"
        )

    print(f"LOADING: Loading cost sheet template: {template_path.name}")
    wb = load_workbook(template_path)
    ws = wb.active

    header_row = 12
    excel_headers = {}
    for idx, cell in enumerate(ws[header_row]):
        val = cell.value
        if isinstance(val, str):
            # Clean up headers by normalizing spaces and converting to uppercase
            header = val.strip().upper().replace('\n', ' ').replace('\xa0', ' ')
            # Replace multiple spaces with single spaces
            import re
            header = re.sub(r'\s+', ' ', header)
            excel_headers[header] = idx + 1

    print("Mapped headers from cost sheet:", excel_headers)

    # Insert data row by row
    for i, row in df_out.iterrows():
        for col_name, value in row.items():
            col_upper = col_name.upper().strip()
            if col_upper in excel_headers:
                col_idx = excel_headers[col_upper]
                # Handle different data types appropriately
                if pd.notna(value) and str(value).strip() != '':
                    ws.cell(row=header_row + 1 + i, column=col_idx, value=value)
                    print(f"📝 Inserted '{value}' into {col_upper} at row {header_row + 1 + i}, col {col_idx}")
                else:
                    # Insert empty string for missing/null values
                    ws.cell(row=header_row + 1 + i, column=col_idx, value="")
            else:
                print(f"⚠️ Column '{col_upper}' not found in cost sheet template")

    print(f"✅ Inserted {len(df_out)} rows of data into cost sheet")

    # Save to PDF directory instead of using output_directory parameter
    # Get the original PDF path from the oem_path (which is based on the merged file)
    oem_path_obj = Path(oem_path)
    # Extract the base name without the suffixes to get back to the original PDF name
    base_name = oem_path_obj.stem.replace("_merged_with_prices", "").replace("_merged", "")
    cost_sheet_path = oem_path_obj.parent / f"{base_name}_cost_sheet.xlsx"
    
    print(f"📁 Saving cost sheet to: {cost_sheet_path}")
    wb.save(cost_sheet_path)
    print(f"Final cost sheet saved to: {cost_sheet_path}")

def main():
    """Main function to be called by the pipeline."""
    oem_path = os.environ.get("OEM_INPUT_PATH")
    merged_path = os.environ.get("MERGED_BOM_PATH")
    map_and_insert_data(oem_path, merged_path)

if __name__ == "__main__":
    main()