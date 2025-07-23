"""
Table Selector GUI for BoMination

This module provides a GUI interface for users to select which extracted tables
they want to include in the final output. It displays all extracted tables with
preview capabilities and allows users to select/deselect tables.
"""

import tkinter as tk
from tkinter import ttk, messagebox
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
import pandas as pd
from pandastable import Table


def show_table_selector(tables):
    """
    Display a GUI window for selecting which tables to include in the final output.
    
    Args:
        tables: List of pandas DataFrames to display for selection
        
    Returns:
        List of selected pandas DataFrames
    """
    selected = []

    def on_submit():
        print(f"DEBUG: ===== CHECKBOX VALIDATION STARTED =====")
        print(f"DEBUG: var_list has {len(var_list)} variables")
        print(f"DEBUG: tables has {len(tables)} tables")
        print(f"DEBUG: Root window exists: {root.winfo_exists()}")
        
        selected_count = 0
        for i, var in enumerate(var_list):
            try:
                table_status = "EMPTY" if tables[i].empty else f"{tables[i].shape[0]}x{tables[i].shape[1]}"
                is_selected = var.get()
                print(f"DEBUG: Table {i+1} ({table_status}): checkbox value = {is_selected} (type: {type(is_selected)})")
                if is_selected:
                    selected.append(tables[i])
                    selected_count += 1
            except Exception as e:
                print(f"DEBUG: ERROR reading checkbox {i+1}: {e}")
                print(f"DEBUG: Variable type: {type(var)}")
                print(f"DEBUG: Variable master: {getattr(var, 'master', 'No master')}")
        
        print(f"DEBUG: Final selection count: {selected_count}")
        print(f"DEBUG: Selected list length: {len(selected)}")
        print(f"DEBUG: ===== CHECKBOX VALIDATION COMPLETED =====")
        
        # Validate that at least one table is selected
        if selected_count == 0:
            messagebox.showwarning(
                "No Tables Selected", 
                "Please select at least one table to continue.\n\n"
                "Use the checkboxes to select the tables you want to include in the output."
            )
            return  # Don't close the window
        
        # ADDITIONAL DEBUG: Log detailed info about selected tables
        print(f"DEBUG: ===== SELECTED TABLES SUMMARY =====")
        for i, table in enumerate(selected):
            print(f"DEBUG: Selected table {i+1}: shape={table.shape}, first_row={table.iloc[0].to_dict() if len(table) > 0 else 'EMPTY'}")
        print(f"DEBUG: ===== END SUMMARY =====")
        
        root.destroy()

    # Try to get existing root window, create new one if none exists
    try:
        # Check if there's already a Tk root window
        import tkinter as tk_module
        existing_root = tk_module._default_root
        if existing_root and existing_root.winfo_exists():
            # Use Toplevel to avoid conflicts with existing Tk instance
            root = ttk.Toplevel(existing_root)
            print("TABLE SELECTOR: Using Toplevel window (existing Tk root found)")
        else:
            raise Exception("No existing root")
    except:
        # Create new root window if none exists - use ttkbootstrap for dark theme
        root = ttk.Window(themename="darkly")
        print("TABLE SELECTOR: Created new ttkbootstrap window with dark theme")
    
    root.title("Select Tables to Keep")
    # Maximize window (zoomed) to show minimize/maximize/close buttons
    root.state('zoomed')
    # Bind Escape key to minimize window for easy exit
    root.bind('<Escape>', lambda e: root.iconify())

    # Create main frame using ttkbootstrap
    main_frame = ttk.Frame(root)
    main_frame.grid(row=0, column=0, sticky='nsew', padx=5, pady=5)
    root.grid_rowconfigure(0, weight=1)
    root.grid_columnconfigure(0, weight=1)
    main_frame.grid_columnconfigure(0, weight=1)

    # Add instructions at the top using ttkbootstrap components
    instructions_frame = ttk.Frame(main_frame)
    instructions_frame.grid(row=0, column=0, sticky='ew', pady=(0, 10))

    title_label = ttk.Label(
        instructions_frame, 
        text="Table Selection", 
        font=('Segoe UI', 16, 'bold'),
        bootstyle="primary"
    )
    title_label.grid(row=0, column=0, sticky='w')
    
    instructions_label = ttk.Label(
        instructions_frame, 
        text="Please review the tables below and CHECK the ones you want to include in the final output.\n"
             "All tables start UNSELECTED - you must check the boxes for tables you want to keep.",
        font=('Segoe UI', 10),
        bootstyle="info",
        justify=tk.LEFT
    )
    instructions_label.grid(row=1, column=0, sticky='w', pady=(5, 10))
    
    # Control buttons frame
    control_frame = ttk.Frame(main_frame)
    control_frame.grid(row=1, column=0, sticky='ew', pady=(0, 10))

    # Create a canvas and scrollbar for the table area using ttkbootstrap
    canvas = tk.Canvas(main_frame)
    scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
    scrollable_frame = ttk.Frame(canvas)

    def on_frame_configure(event):
        """Update scrollregion when frame content changes"""
        canvas.configure(scrollregion=canvas.bbox("all"))

    def on_canvas_configure(event):
        """Update the canvas window width when canvas is resized"""
        # Make the canvas window the same width as the canvas
        canvas_width = event.width
        canvas.itemconfig(canvas_window, width=canvas_width)

    scrollable_frame.bind("<Configure>", on_frame_configure)
    canvas.bind('<Configure>', on_canvas_configure)

    canvas_window = canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    # Add mouse wheel scrolling support
    def _on_mousewheel(event):
        canvas.yview_scroll(int(-1*(event.delta/120)), "units")
    
    def _bind_mousewheel(event):
        canvas.bind_all("<MouseWheel>", _on_mousewheel)
    
    def _unbind_mousewheel(event):
        canvas.unbind_all("<MouseWheel>")
    
    canvas.bind('<Enter>', _bind_mousewheel)
    canvas.bind('<Leave>', _unbind_mousewheel)

    # Grid the canvas and scrollbar
    canvas.grid(row=2, column=0, sticky="nsew", pady=(0, 10))
    scrollbar.grid(row=2, column=1, sticky="ns", pady=(0, 10))
    
    # Configure grid weights for the main frame
    main_frame.grid_rowconfigure(2, weight=1)
    scrollable_frame.grid_columnconfigure(0, weight=1)

    var_list = []
    for i, table in enumerate(tables):
        # Create a frame for each table in the scrollable frame using ttkbootstrap
        frame = ttk.LabelFrame(
            scrollable_frame, 
            text=f"Table {i+1}", 
            padding=10,
            bootstyle="primary"
        )
        frame.grid(row=i, column=0, sticky='ew', padx=5, pady=5)
        scrollable_frame.grid_rowconfigure(i, weight=0)

        # Configure the frame to expand horizontally
        frame.grid_columnconfigure(0, weight=1)

        # Handle empty tables but still add a checkbox variable to maintain index alignment
        if table.empty:
            # Create a simple label for empty table with dark theme styling
            empty_frame = ttk.Frame(frame, bootstyle="secondary")
            empty_frame.pack(fill='both', expand=True, padx=5, pady=5)
            
            empty_label = ttk.Label(
                empty_frame, 
                text="Empty Table", 
                font=('Segoe UI', 12),
                bootstyle="secondary"
            )
            empty_label.pack(expand=True, pady=20)
            
            # Add info about the empty table
            info_label = ttk.Label(
                frame, 
                text="Rows: 0, Columns: 0", 
                font=('Segoe UI', 9),
                bootstyle="secondary"
            )
            info_label.pack(anchor="w", padx=5, pady=2)
            
            # Still add a checkbox variable for empty tables to maintain index alignment
            var = tk.BooleanVar(master=root, value=False)
            # Create a simple frame for the checkbox
            checkbox_frame = ttk.Frame(frame)
            checkbox_frame.pack(anchor="w", padx=5, pady=5)
            cb = ttk.Checkbutton(
                checkbox_frame, 
                text="Select this table (empty)", 
                variable=var, 
                bootstyle="secondary",
                state='disabled'
            )
            cb.pack(anchor="w")
            var_list.append(var)
            continue

        # Store original columns and set up display parameters
        orig_columns = list(table.columns)
        max_rows_to_show = 50  # Limit displayed rows for performance
        
        # Truncate table for display if it's too large
        display_table = table.head(max_rows_to_show) if len(table) > max_rows_to_show else table
        
        # Create a frame for the table with proper configuration using pandastable
        table_frame = ttk.Frame(frame, height=600)
        table_frame.pack(fill='both', expand=True, padx=2, pady=2)
        table_frame.pack_propagate(False)  # Don't shrink the frame
        
        # Configure table_frame grid weights
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)

        # Use pandastable instead of ttk.Treeview for much better display
        try:
            # Ensure we're working with a proper dataframe
            if display_table.empty:
                raise Exception("Display table is empty")
            
            # Make sure the table_frame is properly configured
            table_frame.update_idletasks()  # Ensure frame is ready
            
            # Calculate optimal column widths based on content
            def calculate_optimal_width(series, min_width=150, max_width=400):
                """Calculate optimal column width based on content."""
                if len(series) == 0:
                    return min_width
                
                # Get max length of content in the series
                max_content_length = max(len(str(val)) for val in series if pd.notna(val))
                
                # Calculate width (roughly 10 pixels per character for better spacing)
                calculated_width = max_content_length * 10
                
                # Clamp to min/max bounds
                return max(min_width, min(calculated_width, max_width))
            
            # Calculate column widths for each column
            column_widths = {}
            for col in display_table.columns:
                # Include column name in width calculation with larger multiplier
                col_name_width = len(str(col)) * 10
                content_width = calculate_optimal_width(display_table[col])
                # Use a more generous minimum width
                column_widths[col] = max(col_name_width, content_width, 180)  # Minimum 180px
            
            print(f"TABLE DISPLAY: Column widths calculated: {column_widths}")
            
            # Configure pandastable options for great visibility
            # Simplified options to avoid Tkinter image conflicts
            table_options = {
                'cellwidth': 200,  # Much wider default cell width
                'rowheight': 35,   # Taller row height for better readability 
                'editable': False, # Read-only for selection
                'showstatusbar': False,  # Disable to avoid image conflicts
                'showtoolbar': False,    # Disable to avoid image conflicts
                'font': ('Arial', 11),   # Larger font
                'fontsize': 11,          # Explicit font size
                'align': 'w',            # Left align using proper Tkinter anchor (west)
                'colheadercolor': '#f0f0f0',  # Light header background
                'cellbackgr': 'white',        # White cell background
                'grid_color': '#d0d0d0',      # Visible grid lines
                'linewidth': 1,               # Grid line width
                'showindex': False,           # Disable dataframe index
                'showrowheader': False,       # CRITICAL: Hide row header to eliminate grey space
                'x_start': 0,                 # Set left margin to 0
            }
            
            print(f"TABLE DISPLAY: About to create pandastable with options: {table_options}")
            print(f"TABLE DISPLAY: Table frame exists: {table_frame.winfo_exists()}")
            print(f"TABLE DISPLAY: Root window: {root}")
            print(f"TABLE DISPLAY: Display table shape: {display_table.shape}")
            print(f"TABLE DISPLAY: Display table columns: {list(display_table.columns)}")
            
            # Create table with custom options - use try-catch for safer creation
            try:
                pt = Table(table_frame, dataframe=display_table, **table_options)
                pt.show()
                
                # CRITICAL: Hide row header after show() to eliminate grey space
                try:
                    pt.hideRowHeader()
                    
                    # ADDITIONAL FIX: Reconfigure grid layout to remove column 0 space
                    # Move colheader and main table to column 0 to eliminate left margin
                    pt.colheader.grid(row=0, column=0, rowspan=1, sticky='news')
                    pt.grid(row=1, column=0, rowspan=1, sticky='news', pady=0, ipady=0)
                    
                    # Update column configuration to remove column 0 weight
                    pt.parentframe.columnconfigure(0, weight=1)  # Column 0 gets the weight
                    pt.parentframe.columnconfigure(1, weight=0)  # Column 1 gets no weight
                    
                    print(f"TABLE DISPLAY: Successfully hid row header and reconfigured grid layout")
                except Exception as hide_error:
                    print(f"TABLE DISPLAY: Could not hide row header: {hide_error}")
                
                print(f"TABLE DISPLAY: Created pandastable successfully - shape: {display_table.shape}")
            except Exception as table_error:
                print(f"TABLE DISPLAY: Error creating pandastable: {table_error}")
                # Try with absolute minimal options as fallback
                minimal_options = {
                    'cellwidth': 200,    # Wide cells
                    'rowheight': 35,     # Tall rows
                    'editable': False,
                    'showstatusbar': False,
                    'showtoolbar': False,
                    'font': ('Arial', 11),
                    'fontsize': 11,
                    'showindex': False,   # Disable dataframe index
                    'showrowheader': False,  # CRITICAL: Hide row header to eliminate grey space
                    'x_start': 0,         # Set left margin to 0
                }
                print(f"TABLE DISPLAY: Trying with minimal options: {minimal_options}")
                pt = Table(table_frame, dataframe=display_table, **minimal_options)
                pt.show()
                
                # CRITICAL: Hide row header after show() to eliminate grey space
                try:
                    pt.hideRowHeader()
                    
                    # ADDITIONAL FIX: Reconfigure grid layout to remove column 0 space
                    # Move colheader and main table to column 0 to eliminate left margin
                    pt.colheader.grid(row=0, column=0, rowspan=1, sticky='news')
                    pt.grid(row=1, column=0, rowspan=1, sticky='news', pady=0, ipady=0)
                    
                    # Update column configuration to remove column 0 weight
                    pt.parentframe.columnconfigure(0, weight=1)  # Column 0 gets the weight
                    pt.parentframe.columnconfigure(1, weight=0)  # Column 1 gets no weight
                    
                    print(f"TABLE DISPLAY: Successfully hid row header and reconfigured grid layout (minimal options)")
                except Exception as hide_error:
                    print(f"TABLE DISPLAY: Could not hide row header (minimal options): {hide_error}")
                
                print(f"TABLE DISPLAY: Created pandastable with minimal options")
            
            # Apply custom column widths after table is created
            try:
                # Check if the pandastable version supports columnwidths
                if hasattr(pt.model, 'columnwidths'):
                    # Access the table model and set column widths
                    for col, width in column_widths.items():
                        if col in pt.model.columnwidths:
                            pt.model.columnwidths[col] = width
                        else:
                            # If column not in columnwidths dict, add it
                            pt.model.columnwidths[col] = width
                    
                    # Force redraw to apply the new column widths
                    pt.redraw()
                    print(f"TABLE DISPLAY: Applied custom column widths")
                else:
                    print(f"TABLE DISPLAY: This pandastable version doesn't support column width adjustment")
                    # Try alternative method if available
                    if hasattr(pt, 'adjustColumnWidths'):
                        pt.adjustColumnWidths()
                        print(f"TABLE DISPLAY: Used adjustColumnWidths method")

            except Exception as width_error:
                print(f"TABLE DISPLAY: Could not apply custom column widths: {width_error}")
                # Try fallback method
                try:
                    if hasattr(pt.model, 'columnwidths'):
                        # Set all columns to a wider default
                        for col in display_table.columns:
                            pt.model.columnwidths[col] = 250  # Increased from 200
                        pt.redraw()
                        print(f"TABLE DISPLAY: Applied uniform wider column widths")
                    else:
                        print(f"TABLE DISPLAY: Column width adjustment not supported in this pandastable version")
                except Exception as fallback_error:
                    print(f"TABLE DISPLAY: Could not apply fallback widths: {fallback_error}")
            
        except Exception as e:
            print(f"TABLE DISPLAY: Could not create pandastable: {e}")
            print(f"TABLE DISPLAY: Creating fallback display...")
            
            # Try a simple Treeview as secondary fallback
            try:
                print(f"TABLE DISPLAY: Attempting simple Treeview fallback...")
                
                # Create a simple frame for Treeview
                tree_frame = tk.Frame(table_frame, bg='white')
                tree_frame.pack(fill='both', expand=True)
                
                # Create Treeview with minimal styling
                columns = list(display_table.columns)
                tree = ttk.Treeview(tree_frame, columns=columns, show='headings', height=15)
                
                # Configure column headings with better spacing
                for col in columns:
                    tree.heading(col, text=str(col))
                    # Use wider minimum column widths
                    min_width = max(len(str(col)) * 10, 120)  # At least 120px
                    tree.column(col, width=min_width, minwidth=min_width)
                
                # Set row height for better readability
                tree.configure(style='Treeview')
                style = ttk.Style()
                style.configure('Treeview', rowheight=30)  # Taller rows
                
                # Insert data (limit to first 100 rows for performance)
                max_display_rows = min(100, len(display_table))
                for i in range(max_display_rows):
                    row_data = [str(display_table.iloc[i, j]) for j in range(len(columns))]
                    tree.insert('', 'end', values=row_data)
                
                # Add scrollbars
                v_scrollbar = ttk.Scrollbar(tree_frame, orient='vertical', command=tree.yview)
                h_scrollbar = ttk.Scrollbar(tree_frame, orient='horizontal', command=tree.xview)
                tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
                
                # Layout
                tree.grid(row=0, column=0, sticky='nsew')
                v_scrollbar.grid(row=0, column=1, sticky='ns')
                h_scrollbar.grid(row=1, column=0, sticky='ew')
                
                tree_frame.grid_rowconfigure(0, weight=1)
                tree_frame.grid_columnconfigure(0, weight=1)
                
                if max_display_rows < len(display_table):
                    truncate_label = tk.Label(frame, 
                                            text=f"Showing first {max_display_rows} of {len(display_table)} rows",
                                            font=("Arial", 9), fg='orange', bg='white')
                    truncate_label.pack(pady=(5, 0))
                
                print(f"TABLE DISPLAY: Created Treeview fallback successfully")
                
            except Exception as tree_error:
                print(f"TABLE DISPLAY: Treeview fallback also failed: {tree_error}")
                # Final fallback to text display
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
                content = display_table.to_string(index=False, max_cols=None, max_rows=50)  # Limit rows for display
                text_widget.insert('1.0', content)
                text_widget.configure(state='disabled')
                
                # Add a label to show this is text fallback
                fallback_label = ttk.Label(
                    frame, 
                    text="[WARNING] Using text display (table widgets not available)",
                    font=("Segoe UI", 9), 
                    bootstyle="warning"
                )
                fallback_label.pack(pady=(5, 0))
                
                print(f"🔧 TABLE DISPLAY: Created text fallback successfully")
        
        # Add table info display for non-empty tables
        if not table.empty:
            info_text = f"Rows: {table.shape[0]}, Columns: {table.shape[1]}"
            if len(table) > max_rows_to_show:
                info_text += f" (showing first {max_rows_to_show} rows)"
            
            info_label = ttk.Label(
                frame, 
                text=info_text, 
                font=('Segoe UI', 9),
                bootstyle="secondary"
            )
            info_label.pack(anchor="w", padx=5, pady=2)
        
        # Checkbox to select/deselect table with ttkbootstrap styling
        var = tk.BooleanVar(master=root, value=False)
        
        # Add a callback to test checkbox functionality
        def checkbox_callback():
            print(f"DEBUG: Checkbox {len(var_list)} clicked! New value: {var.get()}")
        
        # Create a simple frame for the checkbox
        checkbox_frame = ttk.Frame(frame)
        checkbox_frame.pack(anchor="w", padx=5, pady=5)
        cb = ttk.Checkbutton(
            checkbox_frame, 
            text="Select this table", 
            variable=var, 
            bootstyle="success",
            command=checkbox_callback
        )
        cb.pack(anchor="w")
        var_list.append(var)

    # NOW define the control functions after var_list is populated
    def select_all():
        """Select all tables"""
        print(f"DEBUG: select_all() called - var_list has {len(var_list)} variables")
        for i, var in enumerate(var_list):
            try:
                var.set(True)
                print(f"DEBUG: Set checkbox {i+1} to True")
            except Exception as e:
                print(f"DEBUG: Error setting checkbox {i+1}: {e}")
    
    def clear_all():
        """Clear all table selections"""
        print(f"DEBUG: clear_all() called - var_list has {len(var_list)} variables")
        for i, var in enumerate(var_list):
            try:
                var.set(False)
                print(f"DEBUG: Set checkbox {i+1} to False")
            except Exception as e:
                print(f"DEBUG: Error clearing checkbox {i+1}: {e}")
    
    # Add the buttons now that the functions are properly defined
    select_all_btn = ttk.Button(
        control_frame, 
        text="Select All", 
        command=select_all, 
        bootstyle="success",
        width=12
    )
    select_all_btn.grid(row=0, column=0, padx=5)
    
    clear_all_btn = ttk.Button(
        control_frame, 
        text="Clear All", 
        command=clear_all, 
        bootstyle="danger",
        width=12
    )
    clear_all_btn.grid(row=0, column=1, padx=5)

    # Continue button
    button_frame = ttk.Frame(main_frame)
    button_frame.grid(row=3, column=0, sticky='ew', padx=5, pady=10)
    
    continue_button = ttk.Button(
        button_frame, 
        text="Continue with Selected Tables", 
        command=on_submit,
        bootstyle="primary",
        width=30
    )
    continue_button.grid(row=0, column=0)
    
    # Add a label showing current selection count
    selection_label = ttk.Label(
        button_frame, 
        text="No tables selected", 
        bootstyle="danger",
        font=('Segoe UI', 10)
    )
    selection_label.grid(row=1, column=0, pady=(10, 0))
    
    def update_selection_count():
        """Update the selection count display"""
        try:
            count = sum(var.get() for var in var_list)
            if count == 0:
                selection_label.config(text="[WARNING] No tables selected", bootstyle="danger")
            elif count == 1:
                selection_label.config(text="[OK] 1 table selected", bootstyle="success")
            else:
                selection_label.config(text=f"[OK] {count} tables selected", bootstyle="success")
            
            # Debug output every 10th update to avoid spam
            if hasattr(update_selection_count, 'debug_counter'):
                update_selection_count.debug_counter += 1
            else:
                update_selection_count.debug_counter = 1
            
            if update_selection_count.debug_counter % 10 == 0:
                print(f"DEBUG: Selection count update #{update_selection_count.debug_counter}: {count} selected")
            
        except Exception as e:
            print(f"DEBUG: Error in update_selection_count: {e}")
            selection_label.config(text="[WARNING] Error reading selections", bootstyle="danger")

        # Schedule next update
        root.after(500, update_selection_count)
    
    # Start the selection count updates
    update_selection_count()
    
    # Instead of running our own mainloop, we'll use wait_window() which is thread-safe
    # and compatible with existing Tkinter applications running in background threads
    print(f"DEBUG: About to start wait_window() for table selector")
    root.wait_window()  # This will block until the window is destroyed
    print(f"DEBUG: wait_window() completed, window destroyed")
    return selected