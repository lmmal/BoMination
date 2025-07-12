#!/usr/bin/env python3
"""
Test script to verify the review functionality works correctly.
This tests the new GUI review callback approach.
"""

import sys
import os
from pathlib import Path

# Add src to path
sys.path.insert(0, str(Path(__file__).parent / "src"))

import pandas as pd
import tkinter as tk
from main_pipeline import run_main_pipeline_with_gui_review
from BoMinationApp import BoMinationApp

def test_review_callback():
    """Test the review callback functionality."""
    print("=== Testing Review Callback ===")
    
    # Create a simple test dataframe
    test_df = pd.DataFrame({
        'Item': ['1', '2', '3'],
        'Quantity': ['10', '5', '8'],
        'Part Number': ['ABC123', 'DEF456', 'GHI789'],
        'Description': ['Test Part 1', 'Test Part 2', 'Test Part 3']
    })
    
    print(f"Test dataframe shape: {test_df.shape}")
    print(f"Test dataframe:\n{test_df}")
    
    # Create a minimal GUI for testing
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    
    # Create the app instance
    app = BoMinationApp(root)
    
    # Test the review window directly
    print("\nTesting review window...")
    try:
        reviewed_df = app.show_review_window(test_df)
        print(f"Review completed! Reviewed dataframe shape: {reviewed_df.shape}")
        print(f"Reviewed dataframe:\n{reviewed_df}")
        print("✓ Review window test passed!")
    except Exception as e:
        print(f"❌ Review window test failed: {e}")
        import traceback
        traceback.print_exc()
    
    root.destroy()

def test_pipeline_with_nel_pdf():
    """Test the full pipeline with NEL PDF."""
    print("\n=== Testing Full Pipeline with NEL PDF ===")
    
    # Use the NEL test PDF
    pdf_path = Path(__file__).parent / "54-0401-0182_rev_c.pdf"
    
    if not pdf_path.exists():
        print(f"❌ Test PDF not found: {pdf_path}")
        return
    
    print(f"Testing with PDF: {pdf_path}")
    
    # Create a minimal GUI for testing
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    
    # Create the app instance
    app = BoMinationApp(root)
    
    # Define a test review callback
    def test_review_callback(merged_df):
        print(f"📝 TEST: Review callback called with dataframe shape: {merged_df.shape}")
        print(f"📝 TEST: Columns: {merged_df.columns.tolist()}")
        print(f"📝 TEST: First few rows:\n{merged_df.head()}")
        
        # For testing, we'll just return the dataframe unchanged
        # In a real scenario, this would call app.show_review_window(merged_df)
        return merged_df
    
    # Test the pipeline
    try:
        result = run_main_pipeline_with_gui_review(
            pdf_path=str(pdf_path),
            pages="all",
            company="NEL",
            output_directory="",
            review_callback=test_review_callback
        )
        
        print(f"✓ Pipeline test completed!")
        print(f"Result: {result}")
        
    except Exception as e:
        print(f"❌ Pipeline test failed: {e}")
        import traceback
        traceback.print_exc()
    
    root.destroy()

if __name__ == "__main__":
    print("BoMination Review Functionality Test")
    print("=" * 40)
    
    # Test 1: Review callback functionality
    test_review_callback()
    
    # Test 2: Full pipeline with NEL PDF
    # test_pipeline_with_nel_pdf()  # Commented out for now to avoid long extraction
    
    print("\n=== All Tests Completed ===")
