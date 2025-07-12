#!/usr/bin/env python3
"""
Test script to verify the updated review window functionality.
"""

import sys
import os
from pathlib import Path
import pandas as pd
import tkinter as tk

# Add src to path
sys.path.insert(0, str(Path(__file__).parent / "src"))

from BoMinationApp import BoMApp

def test_review_window():
    """Test the updated review window with fullscreen and better cell visibility."""
    print("=== Testing Updated Review Window ===")
    
    # Create a test dataframe with realistic BOM data that has long descriptions
    test_df = pd.DataFrame({
        'Item': ['1', '2', '3', '4', '5'],
        'Quantity': ['10', '5', '8', '12', '3'],
        'Part Number': ['ABC-123-456', 'DEF-789-012', 'GHI-345-678', 'JKL-901-234', 'MNO-567-890'],
        'Description': [
            'Cable Assembly, 26 AWG, 4 conductor, PVC jacket, 24 ft length, shielded twisted pair',
            'Connector, Terminal Block, 5.08mm pitch, 12 position, PCB mount, wire-to-board',
            'Ferrite Bead, 1000 ohm @ 100MHz, 3A current rating, 0805 package, EMI suppression',
            'Microcontroller, ARM Cortex-M4, 168MHz, 1MB Flash, 192KB RAM, LQFP100 package',
            'LED Indicator, Red, 3mm diameter, through-hole mount, 20mA forward current, 2.1V drop'
        ],
        'Manufacturer': ['Alpha Wire', 'Phoenix Contact', 'Murata', 'STMicroelectronics', 'Kingbright'],
        'MPN': ['5454C SL005', '1757242', 'BLM21PG221SN1D', 'STM32F407VGT6', 'WP7113ID'],
        'Proton P/N': ['08-1124-000', '08-1125-000', '08-1126-000', '08-1127-000', '08-1128-000'],
        'Notes': ['MOQ 250', 'RoHS compliant', 'Lead-free', 'Extended temperature range', 'High brightness']
    })
    
    print(f"Test dataframe shape: {test_df.shape}")
    print(f"Test dataframe columns: {test_df.columns.tolist()}")
    print(f"Sample data:\n{test_df.head(3)}")
    
    # Create a minimal GUI for testing
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    
    # Create the app instance
    app = BoMApp(root)
    
    # Test the review window
    print("\n🚀 Opening review window...")
    print("   - Window should open in fullscreen mode")
    print("   - Columns should auto-size to fit content")
    print("   - Cell contents should be clearly visible without being cut off")
    print("   - Press ESC to exit fullscreen")
    print("   - Use the 'Auto-Resize Columns' button to optimize column widths")
    print("   - Use the 'Toggle Fullscreen' button to switch modes")
    
    try:
        reviewed_df = app.show_review_window(test_df)
        print(f"\n✅ Review completed successfully!")
        print(f"   - Reviewed dataframe shape: {reviewed_df.shape}")
        print(f"   - Data unchanged: {test_df.equals(reviewed_df)}")
    except Exception as e:
        print(f"❌ Review window test failed: {e}")
        import traceback
        traceback.print_exc()
    
    root.destroy()
    print("\n=== Review Window Test Completed ===")

if __name__ == "__main__":
    print("BoMination Review Window Test")
    print("=" * 40)
    
    test_review_window()
