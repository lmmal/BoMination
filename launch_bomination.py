#!/usr/bin/env python3
"""
Simple test launcher for BoMination application
"""
import sys
import os
from pathlib import Path

# Add the src directory to the Python path
script_dir = Path(__file__).parent
src_dir = script_dir / "src"
sys.path.insert(0, str(src_dir))

def main():
    """Test the main application import and basic functionality."""
    print("Testing BoMination application import...")
    
    try:
        # Test importing the main GUI application
        from gui.BoMinationApp import BoMApp
        print("✅ Successfully imported BoMApp")
        
        # Test importing pipeline components
        from pipeline.main_pipeline import run_main_pipeline_direct
        print("✅ Successfully imported main_pipeline")
        
        # Test importing validation utilities
        from pipeline.validation_utils import validate_pdf_file
        print("✅ Successfully imported validation_utils")
        
        # Test importing customer formatters
        from omni_cust.customer_formatters import apply_customer_formatter
        print("✅ Successfully imported customer_formatters")
        
        print("\n🎉 All imports successful! The application should work correctly.")
        
        # Optionally, you can uncomment the lines below to actually launch the GUI
        # print("\nLaunching GUI...")
        # import tkinter as tk
        # import ttkbootstrap as ttk
        # root = ttk.Window(themename="darkly")
        # app = BoMinationApp(root)
        # root.mainloop()
        
    except Exception as e:
        print(f"❌ Import failed: {e}")
        import traceback
        traceback.print_exc()
        return False
    
    return True

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
