import os
import sys
import subprocess
from pathlib import Path
from validation_utils import validate_page_range, validate_pdf_file, handle_common_errors, generate_output_path
import logging

# Configure logging to reduce noise from third-party libraries
logging.getLogger('selenium').setLevel(logging.WARNING)
logging.getLogger('urllib3').setLevel(logging.WARNING)
logging.getLogger('requests').setLevel(logging.WARNING)

# Support both script and PyInstaller .exe paths
if getattr(sys, 'frozen', False):
    SCRIPT_DIR = Path(sys._MEIPASS) / "src"
else:
    SCRIPT_DIR = Path(__file__).parent

def run_extract_bom():
    print("STEP 1: Extracting BoM tables from PDF...")
    pdf_path = os.environ.get("BOM_PDF_PATH")
    pages = os.environ.get("BOM_PAGE_RANGE")
    company = os.environ.get("BOM_COMPANY")
    output_directory = os.environ.get("BOM_OUTPUT_DIRECTORY")

    env = os.environ.copy()
    env["BOM_PDF_PATH"] = str(pdf_path or "")
    env["BOM_PAGE_RANGE"] = str(pages or "")
    env["BOM_COMPANY"] = str(company or "")
    env["BOM_OUTPUT_DIRECTORY"] = str(output_directory or "")

    command = [sys.executable, str(SCRIPT_DIR / "extract_bom_tab.py")]
    
    print(f"Running command: {' '.join(command)}")
    result = subprocess.run(command, env=env, capture_output=True, text=True)
    
    print(result.stdout)
    if result.returncode != 0:
        print("ERROR: Table extraction failed:")
        print(result.stderr)
        raise subprocess.CalledProcessError(result.returncode, result.args, output=result.stdout, stderr=result.stderr)

    # Generate the expected merged file path - save to PDF directory for debugging
    pdf_dir = Path(pdf_path).parent
    pdf_name = Path(pdf_path).stem
    merged_path = pdf_dir / f"{pdf_name}_merged.xlsx"
    
    print(f"SUCCESS: Expected merged file: {merged_path}")
    return merged_path

def run_price_lookup(merged_path):
    print("STEP 2: Running price lookup...")
    env = os.environ.copy()
    env["BOM_EXCEL_PATH"] = str(merged_path or "")

    command = [sys.executable, str(SCRIPT_DIR / "lookup_price.py")]
    
    print(f"Running command: {' '.join(command)}")
    result = subprocess.run(command, env=env, capture_output=True, text=True)
    
    print(result.stdout)
    if result.returncode != 0:
        print("ERROR: Price lookup failed:")
        print(result.stderr)
        raise subprocess.CalledProcessError(result.returncode, result.args, output=result.stdout, stderr=result.stderr)
    
    # Generate the expected output path for the file with prices - save to PDF directory
    merged_path_obj = Path(merged_path)
    prices_path = merged_path_obj.parent / f"{merged_path_obj.stem}_merged_with_prices.xlsx"
    
    print(f"SUCCESS: Expected prices file: {prices_path}")
    return prices_path

def run_cost_sheet_mapping(prices_path, merged_path):
    print("STEP 3: Mapping to cost sheet template...")
    env = os.environ.copy()
    env["OEM_INPUT_PATH"] = str(prices_path or "")  # File with prices from OEMSecrets
    env["MERGED_BOM_PATH"] = str(merged_path or "")  # Original merged file for descriptions

    command = [sys.executable, str(SCRIPT_DIR / "map_cost_sheet.py")]
    
    print(f"Running command: {' '.join(command)}")
    result = subprocess.run(command, env=env, capture_output=True, text=True)
    
    print(result.stdout)
    if result.returncode != 0:
        print("ERROR: Cost sheet mapping failed:")
        print(result.stderr)
        raise subprocess.CalledProcessError(result.returncode, result.args, output=result.stdout, stderr=result.stderr)

def main():
    """Main pipeline with comprehensive validation and error handling."""
    print("Starting BoMination Pipeline...")
    
    # Validate inputs before starting
    pdf_path = os.environ.get("BOM_PDF_PATH")
    pages = os.environ.get("BOM_PAGE_RANGE") 
    company = os.environ.get("BOM_COMPANY")
    
    print(f"PDF Path: {pdf_path}")
    print(f"Page Range: {pages}")
    print(f"Company: {company or 'None specified'}")
    
    # Input validation
    pdf_valid, pdf_error = validate_pdf_file(pdf_path)
    if not pdf_valid:
        print(f"ERROR: PDF Validation Failed: {pdf_error}")
        sys.exit(1)
    
    pages_valid, pages_error, parsed_ranges = validate_page_range(pages)
    if not pages_valid:
        print(f"ERROR: Page Range Validation Failed: {pages_error}")
        sys.exit(1)
    
    print("SUCCESS: Input validation passed")
    
    try:
        # Step 1: Extract tables from PDF
        print("\n" + "="*50)
        print("STEP 1: EXTRACTING TABLES FROM PDF")
        print("="*50)
        merged_path = run_extract_bom()
        
        if not os.path.exists(merged_path):
            raise Exception(f"Table extraction failed: Expected output file {merged_path} was not created")
        
        print(f"SUCCESS: Step 1 completed successfully")
        
        # Step 2: Price lookup
        print("\n" + "="*50)
        print("STEP 2: RUNNING PRICE LOOKUP")
        print("="*50)
        prices_path = run_price_lookup(merged_path)
        print(f"SUCCESS: Step 2 completed successfully")
        
        # Step 3: Cost sheet mapping
        print("\n" + "="*50)
        print("STEP 3: MAPPING TO COST SHEET")
        print("="*50)
        run_cost_sheet_mapping(prices_path, merged_path)
        print(f"SUCCESS: Step 3 completed successfully")
        
        print("\n" + "="*50)
        print("SUCCESS: BoM AUTOMATION PIPELINE COMPLETED SUCCESSFULLY!")
        print("="*50)
        
    except subprocess.CalledProcessError as e:
        error_msg = handle_common_errors(e.stderr + "\n" + e.stdout if e.stderr else e.stdout)
        print(f"\nERROR: Pipeline step failed:\n{error_msg}")
        print(f"\nTechnical details:")
        print(f"Command: {' '.join(e.cmd)}")
        print(f"Return code: {e.returncode}")
        print(f"Output: {e.stdout}")
        print(f"Errors: {e.stderr}")
        sys.exit(1)
        
    except Exception as e:
        error_msg = handle_common_errors(str(e))
        print(f"\nERROR: Pipeline failed with unexpected error:\n{error_msg}")
        sys.exit(1)

def run_main_pipeline_direct(pdf_path, pages, company, output_directory):
    """
    Direct function call version of the main pipeline for PyInstaller compatibility.
    Sets environment variables and calls the main pipeline functions directly.
    """
    # Import the main functions from other modules
    from extract_bom_tab import main as extract_main
    from lookup_price import main as lookup_main  
    from map_cost_sheet import main as map_main
    
    # Set environment variables for the pipeline steps
    os.environ["BOM_PDF_PATH"] = str(pdf_path)
    os.environ["BOM_PAGE_RANGE"] = str(pages)
    os.environ["BOM_COMPANY"] = str(company)
    os.environ["BOM_OUTPUT_DIRECTORY"] = str(output_directory)
    
    try:
        print("=== BoMination Pipeline Starting ===")
        print(f"PDF: {pdf_path}")
        print(f"Pages: {pages}")
        print(f"Company: {company}")
        print(f"Output: {output_directory}")
        print()
        
        # Step 1: Extract BoM tables
        print("STEP 1: Extracting BoM tables from PDF...")
        extract_main()
        print("✓ BoM extraction completed")
        
        # Generate the expected merged file path after extraction
        # IMPORTANT: Always look for merged files in PDF directory, not output directory
        # The extraction process saves files to PDF directory regardless of output_directory setting
        pdf_dir = Path(pdf_path).parent
        pdf_name = Path(pdf_path).stem
        merged_path = pdf_dir / f"{pdf_name}_merged.xlsx"
        
        print(f"Looking for merged file in PDF directory: {merged_path}")
        print(f"(Output directory setting: {output_directory or 'None - using PDF directory'})")
        
        # Step 2: Lookup prices
        print("\nSTEP 2: Looking up supplier prices...")
        # Set the BOM_EXCEL_PATH for the price lookup step
        os.environ["BOM_EXCEL_PATH"] = str(merged_path)
        
        # Check if merged file exists before price lookup
        if not os.path.exists(merged_path):
            raise Exception(f"Cannot proceed with price lookup: merged file not found at {merged_path}")
        
        print(f"Input file for price lookup: {merged_path}")
        print(f"File exists: {os.path.exists(merged_path)}")
        print(f"File size: {os.path.getsize(merged_path)} bytes")
        
        try:
            print("Attempting to start price lookup process...")
            lookup_main()
            print("✓ Price lookup completed successfully")
        except Exception as lookup_error:
            print(f"❌ Price lookup failed with error: {lookup_error}")
            print("This error prevents price data from being retrieved.")
            print("The pipeline will continue with the merged file, but prices will not be available.")
            
            # Log the full error for debugging
            import traceback
            print(f"Full error traceback:")
            print(traceback.format_exc())
            
            # Provide specific guidance based on error type
            error_str = str(lookup_error).lower()
            if "chromedriver" in error_str or "chrome" in error_str:
                print("\n🔧 ChromeDriver Issue Detected:")
                print("- Ensure Chrome browser is installed and up to date")
                print("- Verify chromedriver.exe is in the src folder")
                print("- Check that antivirus isn't blocking chromedriver.exe")
                print("- Try updating ChromeDriver to match your Chrome version")
            elif "timeout" in error_str:
                print("\n🔧 Timeout Issue Detected:")
                print("- Check your internet connection")
                print("- The website might be temporarily unavailable")
                print("- Try running the pipeline again in a few minutes")
            elif "connection" in error_str:
                print("\n🔧 Connection Issue Detected:")
                print("- Check your internet connection")
                print("- Corporate firewall might be blocking the connection")
                print("- Try running from a different network")
            
            print("\nContinuing without price data...")
        
        # Generate the expected prices file path
        # IMPORTANT: Price lookup saves files to PDF directory, not output directory
        # The lookup_price.py saves files next to the input file (merged_path)
        prices_path = os.path.splitext(merged_path)[0] + "_merged_with_prices.xlsx"
        
        print(f"Looking for prices file: {prices_path}")
        
        # If price lookup failed, use the merged file as fallback
        if not os.path.exists(prices_path):
            print(f"⚠️ Price file not found: {prices_path}")
            print(f"Using merged file as fallback: {merged_path}")
            prices_path = merged_path
        else:
            print(f"✅ Price file found: {prices_path}")
        
        # Step 3: Map to cost sheet
        print("\nSTEP 3: Mapping to OMNI cost sheet template...")
        # Set the environment variables for the cost sheet mapping step
        os.environ["OEM_INPUT_PATH"] = str(prices_path)  # File with prices from OEMSecrets
        os.environ["MERGED_BOM_PATH"] = str(merged_path)  # Original merged file for descriptions
        os.environ["BOM_COMPANY"] = str(company)  # Ensure company parameter is available for mapping
        map_main()
        print("✓ Cost sheet mapping completed")
        
        print("\n=== Pipeline completed successfully! ===")
        return True
        
    except Exception as e:
        error_msg = handle_common_errors(str(e))
        print(f"\nERROR: Pipeline failed:\n{error_msg}")
        raise

def run_main_pipeline_with_gui_review(pdf_path, pages, company, output_directory, review_callback=None):
    """
    Direct function call version of the main pipeline with GUI review callback.
    The review_callback function will be called when the merged table is ready for review.
    """
    # Import the main functions from other modules
    from extract_bom_tab import main as extract_main, merge_tables_and_export
    from lookup_price import main as lookup_main  
    from map_cost_sheet import main as map_main
    
    # Set environment variables for the pipeline steps
    os.environ["BOM_PDF_PATH"] = str(pdf_path)
    os.environ["BOM_PAGE_RANGE"] = str(pages)
    os.environ["BOM_COMPANY"] = str(company)
    os.environ["BOM_OUTPUT_DIRECTORY"] = str(output_directory)
    
    # Skip review in extract_bom_tab.py since we'll handle it in the GUI
    os.environ["BOM_SKIP_REVIEW"] = "true"
    
    try:
        print("=== BoMination Pipeline Starting (GUI Review Mode) ===")
        print(f"PDF: {pdf_path}")
        print(f"Pages: {pages}")
        print(f"Company: {company}")
        print(f"Output: {output_directory}")
        print()
        
        # Step 1: Extract BoM tables
        print("STEP 1: Extracting BoM tables from PDF...")
        extract_main()
        print("✓ BoM extraction completed")
        
        # Generate the expected merged file path after extraction
        pdf_dir = Path(pdf_path).parent
        pdf_name = Path(pdf_path).stem
        merged_path = pdf_dir / f"{pdf_name}_merged.xlsx"
        
        print(f"Looking for merged file: {merged_path}")
        
        if not merged_path.exists():
            raise FileNotFoundError(f"Merged file not found: {merged_path}")
        
        # If we have a review callback, read the merged data and call it
        if review_callback:
            print("Loading merged data for GUI review...")
            import pandas as pd
            merged_df = pd.read_excel(merged_path)
            
            # Call the review callback with the merged data
            print("Calling GUI review callback...")
            reviewed_df = review_callback(merged_df)
            
            # Save the reviewed data back to the file
            if reviewed_df is not None:
                print("Saving reviewed data...")
                reviewed_df.to_excel(merged_path, index=False)
                print("✓ Reviewed data saved")
            else:
                print("No changes made during review")
        
        # Continue with the rest of the pipeline
        print("\nSTEP 2: Looking up part prices...")
        
        # Set the BOM_EXCEL_PATH for the price lookup step
        os.environ["BOM_EXCEL_PATH"] = str(merged_path)
        print(f"Setting BOM_EXCEL_PATH to: {merged_path}")
        
        # Check if merged file exists before price lookup
        if not merged_path.exists():
            raise Exception(f"Cannot proceed with price lookup: merged file not found at {merged_path}")
        
        print(f"Input file for price lookup: {merged_path}")
        print(f"File exists: {merged_path.exists()}")
        print(f"File size: {merged_path.stat().st_size} bytes")
        
        try:
            print("Attempting to start price lookup process...")
            lookup_main()
            print("✓ Price lookup completed successfully")
        except Exception as lookup_error:
            print(f"❌ Price lookup failed with error: {lookup_error}")
            print("This error prevents price data from being retrieved.")
            print("The pipeline will continue with the merged file, but prices will not be available.")
            
            # Log the full error for debugging
            import traceback
            print(f"Full error traceback:")
            print(traceback.format_exc())
        
        # Generate the expected priced file path
        priced_path = pdf_dir / f"{pdf_name}_merged_with_prices.xlsx"
        
        # Check if price lookup was successful
        if not priced_path.exists():
            print(f"⚠️  Warning: Priced file not found at {priced_path}")
            print("Price lookup may have failed, but continuing with cost sheet mapping...")
            print(f"Using merged file as fallback: {merged_path}")
            # Use merged file as fallback for cost sheet mapping
            prices_path = merged_path
        else:
            print(f"✅ Price file found: {priced_path}")
            prices_path = priced_path
        
        print("\nSTEP 3: Mapping to cost sheet...")
        
        # Set the environment variables for the cost sheet mapping step
        os.environ["OEM_INPUT_PATH"] = str(prices_path)  # File with prices (or merged file as fallback)
        os.environ["MERGED_BOM_PATH"] = str(merged_path)  # Original merged file for descriptions
        os.environ["BOM_COMPANY"] = str(company)  # Ensure company parameter is available for mapping
        
        print(f"Setting OEM_INPUT_PATH to: {prices_path}")
        print(f"Setting MERGED_BOM_PATH to: {merged_path}")
        print(f"Setting BOM_COMPANY to: {company}")
        
        map_main()
        print("✓ Cost sheet mapping completed")
        
        # Generate the expected cost sheet file path
        cost_sheet_path = pdf_dir / f"{pdf_name}_cost_sheet.xlsx"
        
        if not cost_sheet_path.exists():
            print(f"⚠️  Warning: Cost sheet not found at {cost_sheet_path}")
            print("Cost sheet mapping may have failed.")
        
        print("\n=== BoMination Pipeline Completed Successfully ===")
        
        # Return information about the generated files
        return {
            "success": True,
            "merged_file": str(merged_path),
            "priced_file": str(prices_path) if prices_path.exists() and prices_path != merged_path else None,
            "cost_sheet_file": str(cost_sheet_path) if cost_sheet_path.exists() else None,
            "pdf_directory": str(pdf_dir)
        }
        
    except Exception as e:
        print(f"❌ Pipeline failed: {e}")
        import traceback
        traceback.print_exc()
        return {"success": False, "error": str(e)}
    
    finally:
        # Clean up environment variable
        if "BOM_SKIP_REVIEW" in os.environ:
            del os.environ["BOM_SKIP_REVIEW"]


if __name__ == "__main__":
    main()