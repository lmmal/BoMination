"""
Test script to verify ROI orchestration architecture
"""
import os
import sys
sys.path.append('.')

print("🧪 Testing ROI orchestration architecture...")

# Test 1: Import the orchestration function
try:
    from pipeline.extract_main import extract_tables_with_roi_orchestration
    print("✅ ROI orchestration function imported successfully")
except Exception as e:
    print(f"❌ Failed to import ROI orchestration: {e}")
    import traceback
    traceback.print_exc()
    sys.exit(1)

# Test 2: Test the orchestration function (dry run)
try:
    # Set up environment variables
    os.environ['BOM_ROI_AREAS'] = '{"1": [100, 50, 400, 300]}'
    
    # Test with a non-existent PDF (should fail gracefully)
    result = extract_tables_with_roi_orchestration("nonexistent.pdf", "1")
    print(f"✅ ROI orchestration function called successfully (result: {len(result)} tables)")
except Exception as e:
    print(f"❌ ROI orchestration failed: {e}")
    import traceback
    traceback.print_exc()

print("🧪 Architecture test completed!")
