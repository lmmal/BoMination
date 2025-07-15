"""
Review Window GUI for BoMination

This module provides a GUI window for reviewing and editing merged BoM tables
before final export. It uses pandastable for editable table display with
fallback options for environments where pandastable is not available.
"""

import tkinter as tk
from tkinter import ttk
import pandas as pd
from pathlib import Path


def show_review_window(merged_df, parent_window=None):
    """
    Show the review window for the merged BoM table.
    
    Args:
        merged_df: pandas DataFrame containing the merged BoM data
        parent_window: Parent tkinter window (optional)
        
    Returns:
        pandas DataFrame with any edits made by the user
    """
    print(f"📝 GUI REVIEW: Creating review window for dataframe with shape {merged_df.shape}")
    
    # Create a new top-level window
    if parent_window:
        review_window = tk.Toplevel(parent_window)
    else:
        review_window = tk.Tk()
    
    review_window.title("Review and Edit Merged BoM Table")
    
    # Make the window fullscreen
    review_window.attributes('-fullscreen', True)
    review_window.configure(bg='white')
    
    # Add escape key to exit fullscreen
    review_window.bind('<Escape>', lambda e: review_window.attributes('-fullscreen', False))
    review_window.bind('<F11>', lambda e: review_window.attributes('-fullscreen', not review_window.attributes('-fullscreen')))
    
    # Make window modal and focused if parent exists
    if parent_window:
        review_window.transient(parent_window)
        review_window.grab_set()
    review_window.focus_set()
    
    # Create main frame
    main_frame = tk.Frame(review_window, padx=20, pady=20, bg='white')
    main_frame.pack(fill='both', expand=True)
    
    # Title with better visibility
    title_label = tk.Label(main_frame, text="Review and Edit Merged BoM Table", 
                          font=("Arial", 20, "bold"), bg='white', fg='darkblue')
    title_label.pack(pady=(0, 15))
    
    # Instructions with better formatting
    instructions = tk.Label(main_frame, 
                           text="📋 Review the merged table below. You can edit values directly in the table.\n"
                                "💾 Click 'Save & Continue' when done, or 'Cancel' to use the data as-is.\n"
                                "⌨️ Press ESC to exit fullscreen mode.",
                           font=("Arial", 12), justify='left', bg='white', fg='darkgreen')
    instructions.pack(pady=(0, 15))
    
    # Add table info
    table_info = tk.Label(main_frame, 
                         text=f"📊 Table: {merged_df.shape[0]} rows × {merged_df.shape[1]} columns | "
                              f"Columns: {', '.join(merged_df.columns.tolist()[:5])}{'...' if len(merged_df.columns) > 5 else ''}",
                         font=("Arial", 10), bg='white', fg='darkslategray')
    table_info.pack(pady=(0, 10))
    
    # Variable to store the result
    result_df = [merged_df.copy()]  # Use list to allow modification in nested functions
    review_completed = [False]  # Flag to track if review was completed
    
    # Create table frame with white background
    table_frame = tk.Frame(main_frame, bg='white', relief='ridge', bd=2)
    table_frame.pack(fill='both', expand=True, pady=(0, 10))
    
    # Try to create an editable table
    table_widget = None
    try:
        from pandastable import Table
        
        # Calculate optimal column widths based on content
        def calculate_optimal_width(series, min_width=80, max_width=400):
            """Calculate optimal column width based on content."""
            if len(series) == 0:
                return min_width
            
            # Get max length of content in the series
            max_content_length = max(len(str(val)) for val in series)
            
            # Calculate width (roughly 8 pixels per character)
            calculated_width = max_content_length * 8
            
            # Clamp to min/max bounds
            return max(min_width, min(calculated_width, max_width))
        
        # Calculate column widths for each column
        column_widths = {}
        for col in merged_df.columns:
            # Include column name in width calculation
            col_name_width = len(str(col)) * 8
            content_width = calculate_optimal_width(merged_df[col])
            column_widths[col] = max(col_name_width, content_width)
        
        print(f"📝 GUI REVIEW: Column widths calculated: {column_widths}")
        
        # Configure pandastable options for better visibility
        table_options = {
            'cellwidth': 200,  # Default wider cells
            'cellbackgr': 'white',
            'grid_color': 'gray',
            'rowheight': 30,  # Taller rows for better readability
            'font': ('Arial', 10),
            'fontsize': 10,
            'align': 'w',  # Left align
            'precision': 2,
            'showstatusbar': True,
            'showtoolbar': True,
            'editable': True,
            'wrap': True,  # Enable text wrapping
            'linewidth': 1,
            'entrybackgr': 'white',
            'rowselectedcolor': '#E6F3FF',
            'multipleselectioncolor': '#CCE7FF',
            'showindex': False,  # Disable dataframe index
            'showrowheader': False,  # CRITICAL: Hide row header to eliminate grey space
            'x_start': 0,        # Set left margin to 0
        }
        
        # Create table with custom options
        table_widget = Table(table_frame, dataframe=merged_df, **table_options)
        table_widget.show()
        
        # CRITICAL: Hide row header after show() to eliminate grey space
        try:
            table_widget.hideRowHeader()
            
            # ADDITIONAL FIX: Reconfigure grid layout to remove column 0 space
            # Move colheader and main table to column 0 to eliminate left margin
            table_widget.colheader.grid(row=0, column=0, rowspan=1, sticky='news')
            table_widget.grid(row=1, column=0, rowspan=1, sticky='news', pady=0, ipady=0)
            
            # Update column configuration to remove column 0 weight
            table_widget.parentframe.columnconfigure(0, weight=1)  # Column 0 gets the weight
            table_widget.parentframe.columnconfigure(1, weight=0)  # Column 1 gets no weight
            
            print(f"📝 GUI REVIEW: Successfully hid row header and reconfigured grid layout")
        except Exception as hide_error:
            print(f"📝 GUI REVIEW: Could not hide row header: {hide_error}")
        
        # Apply custom column widths after table is created
        try:
            # Access the table model and set column widths
            for col, width in column_widths.items():
                if col in table_widget.model.columnwidths:
                    table_widget.model.columnwidths[col] = width
                else:
                    # If column not in columnwidths dict, add it
                    table_widget.model.columnwidths[col] = width
            
            # Force redraw to apply the new column widths
            table_widget.redraw()
            print(f"📝 GUI REVIEW: Applied custom column widths")
            
        except Exception as width_error:
            print(f"📝 GUI REVIEW: Could not apply custom column widths: {width_error}")
            # Fall back to uniform wider columns
            try:
                # Set all columns to a wider default
                for col in merged_df.columns:
                    table_widget.model.columnwidths[col] = 250
                table_widget.redraw()
                print(f"📝 GUI REVIEW: Applied uniform wider column widths")
            except Exception as fallback_error:
                print(f"📝 GUI REVIEW: Could not apply fallback widths: {fallback_error}")
        
        print(f"📝 GUI REVIEW: Created editable table with pandastable")
        print(f"📝 GUI REVIEW: Table dimensions: {merged_df.shape}")
        print(f"📝 GUI REVIEW: Sample data:\n{merged_df.head(2)}")
        
    except Exception as e:
        print(f"📝 GUI REVIEW: Could not create pandastable: {e}")
        print(f"📝 GUI REVIEW: Creating fallback text view...")
        
        # Enhanced fallback to text view with better formatting
        text_widget = tk.Text(table_frame, wrap='none', font=("Courier", 9))
        
        # Add scrollbars
        v_scrollbar = tk.Scrollbar(table_frame, orient='vertical', command=text_widget.yview)
        h_scrollbar = tk.Scrollbar(table_frame, orient='horizontal', command=text_widget.xview)
        text_widget.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        # Layout
        text_widget.grid(row=0, column=0, sticky='nsew')
        v_scrollbar.grid(row=0, column=1, sticky='ns')
        h_scrollbar.grid(row=1, column=0, sticky='ew')
        
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)
        
        # Insert table content with better formatting
        content = merged_df.to_string(index=False, max_cols=None, max_rows=None)
        text_widget.insert('1.0', content)
        text_widget.configure(state='disabled')
        
        # Add a label to show this is read-only
        readonly_label = tk.Label(main_frame, 
                                text="⚠️ Table is read-only (pandastable not available). Data will be used as-is.",
                                font=("Arial", 10), fg='orange')
        readonly_label.pack(pady=(5, 0))
    
    # Button functions
    def save_and_continue():
        print(f"📝 GUI REVIEW: User clicked Save & Continue")
        if table_widget is not None:
            try:
                # Get edited data from table
                edited_data = table_widget.model.df.copy()
                result_df[0] = edited_data
                print(f"📝 GUI REVIEW: Retrieved edited data, shape: {edited_data.shape}")
            except Exception as e:
                print(f"📝 GUI REVIEW: Could not get edited data: {e}")
                result_df[0] = merged_df.copy()
        else:
            result_df[0] = merged_df.copy()
        
        review_completed[0] = True
        review_window.destroy()
    
    def cancel_review():
        print(f"📝 GUI REVIEW: User clicked Cancel")
        result_df[0] = merged_df.copy()
        review_completed[0] = True
        review_window.destroy()
    
    # Create buttons with better visibility
    button_frame = tk.Frame(main_frame, bg='white')
    button_frame.pack(fill='x', pady=(15, 0))
    
    # Auto-resize columns function
    def auto_resize_columns():
        """Auto-resize columns to fit content."""
        if table_widget is not None:
            try:
                # Get current dataframe from table
                current_df = table_widget.model.df
                
                # Recalculate column widths
                for col in current_df.columns:
                    # Calculate optimal width for this column
                    col_name_width = len(str(col)) * 8
                    max_content_length = max(len(str(val)) for val in current_df[col]) if len(current_df[col]) > 0 else 0
                    content_width = max_content_length * 8
                    optimal_width = max(col_name_width, content_width, 100)  # Minimum 100px
                    optimal_width = min(optimal_width, 500)  # Maximum 500px
                    
                    # Apply the width
                    table_widget.model.columnwidths[col] = optimal_width
                
                # Redraw table
                table_widget.redraw()
                print(f"📝 GUI REVIEW: Auto-resized columns")
                
            except Exception as e:
                print(f"📝 GUI REVIEW: Could not auto-resize columns: {e}")
    
    # Exit fullscreen button
    def toggle_fullscreen():
        current_state = review_window.attributes('-fullscreen')
        review_window.attributes('-fullscreen', not current_state)
        if current_state:
            # Exiting fullscreen, set a reasonable size
            review_window.geometry("1400x900")
            # Center it
            review_window.update_idletasks()
            x = (review_window.winfo_screenwidth() // 2) - (1400 // 2)
            y = (review_window.winfo_screenheight() // 2) - (900 // 2)
            review_window.geometry(f"1400x900+{x}+{y}")
    
    # Auto-resize button (only show if table widget exists)
    if table_widget is not None:
        auto_resize_button = tk.Button(button_frame, text="📏 Auto-Resize Columns", command=auto_resize_columns,
                                     bg='lightyellow', font=("Arial", 11), padx=15, pady=5)
        auto_resize_button.pack(side='left', padx=(0, 10))
    
    fullscreen_button = tk.Button(button_frame, text="📺 Toggle Fullscreen", command=toggle_fullscreen,
                                 bg='lightblue', font=("Arial", 11), padx=15, pady=5)
    fullscreen_button.pack(side='left', padx=(0, 10))
    
    save_button = tk.Button(button_frame, text="✅ Save & Continue", command=save_and_continue,
                           bg='lightgreen', font=("Arial", 14, "bold"), padx=25, pady=8)
    save_button.pack(side='left', padx=(0, 10))
    
    cancel_button = tk.Button(button_frame, text="❌ Cancel", command=cancel_review,
                             bg='lightcoral', font=("Arial", 14, "bold"), padx=25, pady=8)
    cancel_button.pack(side='left')
    
    # Handle window close
    def on_window_close():
        print(f"📝 GUI REVIEW: Window closed without explicit action")
        result_df[0] = merged_df.copy()
        review_completed[0] = True
        review_window.destroy()
    
    review_window.protocol("WM_DELETE_WINDOW", on_window_close)
    
    # Wait for the review to complete
    print(f"📝 GUI REVIEW: Waiting for user review...")
    review_window.wait_window()
    
    print(f"📝 GUI REVIEW: Review completed, returning dataframe with shape: {result_df[0].shape}")
    return result_df[0]


def review_and_edit_dataframe_cli(df):
    """
    Review and edit dataframe for command-line usage.
    This is a simpler version for when running outside the main GUI.
    """
    print(f"📝 REVIEW DEBUG: Starting review window for dataframe with shape {df.shape}")
    
    try:
        # Create the window with better configuration
        root = tk.Tk()
        root.title("Review and Edit Merged BoM Table")
        root.geometry("1200x800")
        
        # Force the window to be on top and get focus
        root.lift()
        root.attributes('-topmost', True)
        root.after(100, lambda: root.attributes('-topmost', False))
        
        # Center the window
        root.update_idletasks()
        x = (root.winfo_screenwidth() // 2) - (1200 // 2)
        y = (root.winfo_screenheight() // 2) - (800 // 2)
        root.geometry(f"1200x800+{x}+{y}")
        
        print(f"📝 REVIEW DEBUG: Window created and positioned")
        
        # Main frame
        main_frame = tk.Frame(root)
        main_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Title
        title_label = tk.Label(main_frame, text="Review and Edit Merged BoM Table", 
                              font=("Arial", 14, "bold"))
        title_label.pack(pady=(0, 10))
        
        # Instructions
        instructions = tk.Label(main_frame, 
                               text="Review the merged table below. You can edit cell values directly in the table.\nClick 'Continue' when you're satisfied with the data.",
                               font=("Arial", 10))
        instructions.pack(pady=(0, 10))
        
        # Table frame
        table_frame = tk.Frame(main_frame)
        table_frame.pack(fill='both', expand=True)

        # Try to create the pandastable with error handling
        try:
            from pandastable import Table
            
            print(f"📝 REVIEW DEBUG: Creating pandastable...")
            pt = Table(table_frame, dataframe=df, showtoolbar=True, showstatusbar=True, 
                      showindex=False, x_start=0)  # Disable row index and set left margin to 0
            pt.show()
            
            # Hide the row header to save space
            try:
                pt.hideRowHeader()
                print(f"📝 REVIEW DEBUG: Row header hidden")
            except:
                print(f"📝 REVIEW DEBUG: Could not hide row header")
            
            print(f"📝 REVIEW DEBUG: Pandastable created successfully")
            
            # Variable to hold the final result
            result_df = [df.copy()]
            
            # Button function
            def continue_with_data():
                print(f"📝 REVIEW DEBUG: User clicked continue")
                try:
                    # Get the current dataframe from the table
                    current_df = pt.model.df
                    result_df[0] = current_df.copy()
                    print(f"📝 REVIEW DEBUG: Data retrieved from table, shape: {current_df.shape}")
                except Exception as e:
                    print(f"📝 REVIEW DEBUG: Error getting data from table: {e}")
                    result_df[0] = df.copy()
                
                root.destroy()
            
            # Button frame
            button_frame = tk.Frame(main_frame)
            button_frame.pack(fill='x', pady=(10, 0))
            
            continue_button = tk.Button(button_frame, text="✅ Continue", 
                                       command=continue_with_data,
                                       bg='lightgreen', font=("Arial", 12, "bold"))
            continue_button.pack(side='right')
            
            # Handle window close
            def on_closing():
                print(f"📝 REVIEW DEBUG: Window closed")
                result_df[0] = df.copy()
                root.destroy()
            
            root.protocol("WM_DELETE_WINDOW", on_closing)
            
            print(f"📝 REVIEW DEBUG: Starting mainloop...")
            root.mainloop()
            
            print(f"📝 REVIEW DEBUG: Review completed, returning dataframe with shape: {result_df[0].shape}")
            return result_df[0]
            
        except Exception as e:
            print(f"📝 REVIEW DEBUG: Error creating pandastable: {e}")
            print(f"📝 REVIEW DEBUG: Falling back to text display...")
            
            # Fallback to text view
            text_widget = tk.Text(table_frame, wrap='word', font=("Courier", 9))
            scrollbar = tk.Scrollbar(table_frame, orient='vertical', command=text_widget.yview)
            text_widget.configure(yscrollcommand=scrollbar.set)
            
            text_widget.pack(side='left', fill='both', expand=True)
            scrollbar.pack(side='right', fill='y')
            
            # Insert table content
            content = df.to_string(index=False, max_cols=None, max_rows=None)
            text_widget.insert('1.0', content)
            text_widget.configure(state='disabled')
            
            # Add read-only warning
            warning_label = tk.Label(main_frame, 
                                   text="⚠️ Table is read-only (pandastable not available). Data will be used as-is.",
                                   font=("Arial", 10), fg='orange')
            warning_label.pack(pady=(5, 0))
            
            # Button function for read-only mode
            def continue_readonly():
                print(f"📝 REVIEW DEBUG: User continued with read-only data")
                root.destroy()
            
            # Button frame for read-only mode
            button_frame = tk.Frame(main_frame)
            button_frame.pack(fill='x', pady=(10, 0))
            
            continue_button = tk.Button(button_frame, text="✅ Continue", 
                                       command=continue_readonly,
                                       bg='lightgreen', font=("Arial", 12, "bold"))
            continue_button.pack(side='right')
            
            # Handle window close for read-only mode
            root.protocol("WM_DELETE_WINDOW", continue_readonly)
            
            print(f"📝 REVIEW DEBUG: Starting mainloop (read-only mode)...")
            root.mainloop()
            
            print(f"📝 REVIEW DEBUG: Review completed (read-only), returning original dataframe")
            return df.copy()

    except Exception as e:
        print(f"📝 REVIEW DEBUG: Major error in review window: {e}")
        import traceback
        traceback.print_exc()
        return df.copy()


# For testing the module independently
if __name__ == "__main__":
    import pandas as pd
    
    # Create sample data for testing
    sample_data = {
        'Item': ['Widget A', 'Widget B', 'Widget C'],
        'Quantity': [10, 20, 15],
        'Price': [1.50, 2.75, 3.25],
        'Total': [15.00, 55.00, 48.75]
    }
    
    test_df = pd.DataFrame(sample_data)
    
    print("Testing review window with sample data...")
    result = show_review_window(test_df)
    print(f"Result shape: {result.shape}")
    print(result)
