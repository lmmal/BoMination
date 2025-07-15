#!/usr/bin/env python3
"""
Test script to verify all imports work correctly after directory reorganization.
"""
import sys
import os
from pathlib import Path

# Add src to path
src_path = Path(__file__).parent / "src"
sys.path.insert(0, str(src_path))

def test_imports():
    print("Testing imports after directory reorganization...")
    
    try:
        print("1. Testing pipeline modules...")
        from pipeline.main_pipeline import run_main_pipeline_direct
        from pipeline.extract_bom_tab import extract_tables_with_tabula, merge_tables_and_export
        from pipeline.lookup_price import main
        from pipeline.map_cost_sheet import map_and_insert_data
        from pipeline.validation_utils import validate_pdf_file, validate_page_range
        print("   ✓ All pipeline modules imported successfully")
        
        print("2. Testing GUI modules...")
        from gui.BoMinationApp import BoMinationApp
        from gui.review_window import show_review_window, review_and_edit_dataframe_cli
        from gui.table_selector import show_table_selector
        print("   ✓ All GUI modules imported successfully")
        
        print("3. Testing customer modules...")
        from omni_cust.customer_formatters import apply_customer_formatting
        from omni_cust.customer_config import DEFAULT_CUSTOMER, CUSTOMER_SETTINGS
        print("   ✓ All customer modules imported successfully")
        
        print("\n🎉 All imports working correctly!")
        return True
        
    except ImportError as e:
        print(f"❌ Import error: {e}")
        return False
    except Exception as e:
        print(f"❌ Unexpected error: {e}")
        return False

if __name__ == "__main__":
    success = test_imports()
    sys.exit(0 if success else 1)
