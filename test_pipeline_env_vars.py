#!/usr/bin/env python3
"""
Test script to verify the pipeline environment variable fix.
"""

import sys
import os
from pathlib import Path

# Add src to path
sys.path.insert(0, str(Path(__file__).parent / "src"))

def test_pipeline_env_vars():
    """Test that environment variables are set correctly in the pipeline."""
    print("=== Testing Pipeline Environment Variables ===")
    
    # Test PDF path
    test_pdf = Path(__file__).parent / "54-0401-0182_rev_c.pdf"
    
    if not test_pdf.exists():
        print(f"❌ Test PDF not found: {test_pdf}")
        return
    
    print(f"Using test PDF: {test_pdf}")
    
    # Import the pipeline function
    from main_pipeline import run_main_pipeline_with_gui_review
    
    # Create a mock review callback that just returns the data unchanged
    def mock_review_callback(merged_df):
        print(f"📝 MOCK REVIEW: Called with dataframe shape: {merged_df.shape}")
        print(f"📝 MOCK REVIEW: Columns: {merged_df.columns.tolist()}")
        # Return the data unchanged
        return merged_df
    
    # Test the pipeline with the mock callback
    try:
        print("\n🚀 Testing pipeline with mock review callback...")
        result = run_main_pipeline_with_gui_review(
            pdf_path=str(test_pdf),
            pages="all",
            company="NEL",
            output_directory="",
            review_callback=mock_review_callback
        )
        
        print(f"\n✅ Pipeline test completed!")
        print(f"Result: {result}")
        
        if result.get("success"):
            print("✅ Pipeline reported success")
            print(f"✅ Merged file: {result.get('merged_file')}")
            print(f"✅ Priced file: {result.get('priced_file')}")
            print(f"✅ Cost sheet file: {result.get('cost_sheet_file')}")
        else:
            print("❌ Pipeline reported failure")
            print(f"❌ Error: {result.get('error')}")
        
    except Exception as e:
        print(f"❌ Pipeline test failed: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    print("BoMination Pipeline Environment Variables Test")
    print("=" * 50)
    
    test_pipeline_env_vars()
    
    print("\n=== Test Completed ===")
