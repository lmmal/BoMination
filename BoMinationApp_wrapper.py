#!/usr/bin/env python3
"""
BoMination Application Entry Point
This wrapper script ensures proper module imports for PyInstaller builds.
"""

import sys
import os
from pathlib import Path

# Add the src directory to the Python path
current_dir = Path(__file__).parent
src_dir = current_dir / "src"

# Ensure src is in the path for imports
if str(src_dir) not in sys.path:
    sys.path.insert(0, str(src_dir))

# Now import and run the main application
if __name__ == "__main__":
    try:
        from gui.BoMinationApp import BoMApp
        import ttkbootstrap as ttk
        from pathlib import Path
        import sys
        
        # Create the main window with ttkbootstrap's darkly theme
        root = ttk.Window(themename="darkly")
        
        # Set application icon with improved quality handling
        try:
            if getattr(sys, 'frozen', False):
                # Running as PyInstaller executable
                icon_path = Path(sys._MEIPASS) / "assets" / "BoMination_black.ico"
            else:
                # Running as script - go up one level from root to assets
                icon_path = current_dir / "assets" / "BoMination_black.ico"
            
            if icon_path.exists():
                # Method 1: Use iconbitmap for basic compatibility
                root.iconbitmap(str(icon_path))
                
                # Method 2: Try to set higher quality icon using photoimage
                try:
                    from PIL import Image, ImageTk
                    import tkinter as tk
                    
                    # Load the icon as a PhotoImage for better quality
                    # We'll use a 32x32 version for the window
                    icon_image = Image.open(str(icon_path))
                    icon_image = icon_image.resize((32, 32), Image.Resampling.LANCZOS)
                    photo = ImageTk.PhotoImage(icon_image)
                    root.iconphoto(False, photo)
                    
                    # Store reference to prevent garbage collection
                    root._icon_photo = photo
                    print(f"✅ High-quality icon loaded from: {icon_path}")
                except ImportError:
                    print(f"✅ Basic icon loaded from: {icon_path} (PIL not available for high-quality mode)")
                except Exception as e:
                    print(f"✅ Basic icon loaded from: {icon_path} (high-quality mode failed: {e})")
            else:
                print(f"⚠️ Icon not found at: {icon_path}")
        except Exception as e:
            print(f"⚠️ Could not load application icon: {e}")
        
        root.iconify()  # Start minimized to prevent flash
        
        # Add proper cleanup handler
        def on_closing():
            try:
                root.quit()
                root.destroy()
            except:
                pass  # Ignore cleanup errors
        
        root.protocol("WM_DELETE_WINDOW", on_closing)
        
        app = BoMApp(root)
        root.deiconify()  # Show the window after initialization
        root.mainloop()
        
    except Exception as e:
        print(f"Application error: {e}")
        import traceback
        traceback.print_exc()
    finally:
        # Force cleanup
        try:
            if 'root' in locals():
                root.quit()
                root.destroy()
        except:
            pass
