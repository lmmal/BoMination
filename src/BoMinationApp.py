import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from ttkbootstrap.dialogs import Messagebox
import subprocess
import os
import threading
import sys
import time
import webbrowser
from pathlib import Path
from datetime import datetime
import warnings

# Import the pipeline modules directly for PyInstaller compatibility
try:
    from main_pipeline import run_main_pipeline_direct
except ImportError:
    # Fallback for development mode
    def run_main_pipeline_direct(*args, **kwargs):
        raise ImportError("Pipeline modules not available")

# Suppress known Tkinter destructor warnings in Python 3.12
warnings.filterwarnings("ignore", category=DeprecationWarning, module="tkinter")

# Monkey patch to fix Tkinter Image destructor issue in Python 3.12
try:
    import tkinter
    original_image_del = tkinter.Image.__del__
    
    def safe_image_del(self):
        try:
            original_image_del(self)
        except (TypeError, Exception):
            pass  # Ignore destructor errors
    
    tkinter.Image.__del__ = safe_image_del
except (AttributeError, ImportError):
    pass  # Skip if not available
from validation_utils import (
    validate_page_range, 
    validate_pdf_file, 
    check_java_installation, 
    check_chromedriver_availability,
    handle_common_errors,
    open_help_url,
    validate_output_directory
)

# Default path to cost sheet template
# Support both script and PyInstaller .exe paths
if getattr(sys, 'frozen', False):
    SCRIPT_DIR = Path(sys._MEIPASS) / "src"
    COST_SHEET_TEMPLATE = Path(sys._MEIPASS) / "Files" / "OCTF-1539-COST SHEET.xlsx"
else:
    SCRIPT_DIR = Path(__file__).parent
    COST_SHEET_TEMPLATE = SCRIPT_DIR.parent / "Files" / "OCTF-1539-COST SHEET.xlsx"

# Adding debug logging to trace imports
import logging
logging.basicConfig(level=logging.DEBUG)

# Suppress verbose third-party library logging
logging.getLogger('selenium').setLevel(logging.WARNING)
logging.getLogger('urllib3').setLevel(logging.WARNING)
logging.getLogger('requests').setLevel(logging.WARNING)
logging.getLogger('PIL').setLevel(logging.WARNING)
logging.getLogger('matplotlib').setLevel(logging.WARNING)
logging.getLogger('fontTools').setLevel(logging.WARNING)

try:
    import numpy
    logging.debug("NumPy imported successfully.")
    import PIL
    logging.debug("PIL imported successfully.")
    from ttkbootstrap import Style
    logging.debug("ttkbootstrap imported successfully.")
except Exception as e:
    logging.error(f"Error during imports: {e}")

class CopyableErrorDialog:
    """Custom dialog that allows copying error messages with modern styling."""
    
    def __init__(self, parent, title, message, technical_details=None):
        self.result = None
        
        # Create the dialog window
        self.dialog = tk.Toplevel(parent)
        self.dialog.title(title)
        self.dialog.geometry("700x500")
        self.dialog.resizable(True, True)
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        # Center the dialog
        self.dialog.update_idletasks()
        x = (self.dialog.winfo_screenwidth() // 2) - (700 // 2)
        y = (self.dialog.winfo_screenheight() // 2) - (500 // 2)
        self.dialog.geometry(f"700x500+{x}+{y}")
        
        # Main frame with padding
        main_frame = ttk.Frame(self.dialog)
        
        
        # Title with icon
        title_frame = ttk.Frame(main_frame)
        title_frame.pack(fill=X, pady=(0, 15))
        
        ttk.Label(
            title_frame, 
            text=f"{title}",
            font=("Segoe UI", 16, "bold"),
            bootstyle="danger"
        ).pack(anchor=W)
        
        # Error message section
        msg_frame = ttk.LabelFrame(main_frame, text="Error Message", padding=10)
        msg_frame.pack(fill=BOTH, expand=True, pady=(0, 10))
        
        # Scrollable text widget for the message
        msg_scroll_frame = ttk.Frame(msg_frame)
        msg_scroll_frame.pack(fill=BOTH, expand=True)
        
        msg_text = tk.Text(
            msg_scroll_frame, 
            wrap=tk.WORD, 
            height=8, 
            font=("Segoe UI", 10),
            relief="flat",
            borderwidth=0
        )
        msg_scrollbar = ttk.Scrollbar(msg_scroll_frame, orient=VERTICAL, command=msg_text.yview)
        msg_text.configure(yscrollcommand=msg_scrollbar.set)
        
        msg_text.pack(side=LEFT, fill=BOTH, expand=True)
        msg_scrollbar.pack(side=RIGHT, fill=Y)
        
        msg_text.insert(tk.END, message)
        msg_text.config(state=tk.DISABLED)
        
        # Technical details (if provided)
        if technical_details:
            tech_frame = ttk.LabelFrame(main_frame, text="Technical Details", padding=10)
            tech_frame.pack(fill=BOTH, expand=True, pady=(0, 15))
            
            tech_scroll_frame = ttk.Frame(tech_frame)
            tech_scroll_frame.pack(fill=BOTH, expand=True)
            
            tech_text = tk.Text(
                tech_scroll_frame, 
                wrap=tk.WORD, 
                height=6, 
                font=("Consolas", 9),
                relief="flat",
                borderwidth=0
            )
            tech_scrollbar = ttk.Scrollbar(tech_scroll_frame, orient=VERTICAL, command=tech_text.yview)
            tech_text.configure(yscrollcommand=tech_scrollbar.set)
            
            tech_text.pack(side=LEFT, fill=BOTH, expand=True)
            tech_scrollbar.pack(side=RIGHT, fill=Y)
            
            tech_text.insert(tk.END, technical_details)
            tech_text.config(state=tk.DISABLED)
            
            # Store references for copying
            self.msg_text = msg_text
            self.tech_text = tech_text
        else:
            self.msg_text = msg_text
            self.tech_text = None
        
        # Button frame with modern styling
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=X)
        
        # Left side buttons
        left_buttons = ttk.Frame(button_frame)
        left_buttons.pack(side=LEFT)
        
        ttk.Button(
            left_buttons, 
            text="Copy Error Message", 
            command=self.copy_message,
            bootstyle="primary-outline",
            width=20
        ).pack(side=LEFT, padx=(0, 10))
        
        if technical_details:
            ttk.Button(
                left_buttons, 
                text="Copy Technical Details", 
                command=self.copy_technical,
                bootstyle="secondary-outline",
                width=22
            ).pack(side=LEFT, padx=(0, 10))
            
            ttk.Button(
                left_buttons, 
                text="Copy All", 
                command=self.copy_all,
                bootstyle="info-outline",
                width=12
            ).pack(side=LEFT, padx=(0, 10))
        
        # Close button on the right
        ttk.Button(
            button_frame, 
            text="Close", 
            command=self.close,
            bootstyle="secondary",
            width=12
        ).pack(side=RIGHT)
        
        # Bind escape key to close
        self.dialog.bind('<Escape>', lambda e: self.close())
        
        # Focus on the dialog
        self.dialog.focus_set()
        
    def copy_message(self):
        """Copy the error message to clipboard."""
        self.msg_text.config(state=tk.NORMAL)
        message = self.msg_text.get(1.0, tk.END).strip()
        self.msg_text.config(state=tk.DISABLED)
        
        self.dialog.clipboard_clear()
        self.dialog.clipboard_append(message)
        self.dialog.update()
        
        # Show brief confirmation
        self.show_copy_confirmation("Error message copied to clipboard!")
        
    def copy_technical(self):
        """Copy technical details to clipboard."""
        if self.tech_text:
            self.tech_text.config(state=tk.NORMAL)
            technical = self.tech_text.get(1.0, tk.END).strip()
            self.tech_text.config(state=tk.DISABLED)
            
            self.dialog.clipboard_clear()
            self.dialog.clipboard_append(technical)
            self.dialog.update()
            
            self.show_copy_confirmation("Technical details copied to clipboard!")
    
    def copy_all(self):
        """Copy both error message and technical details."""
        self.msg_text.config(state=tk.NORMAL)
        message = self.msg_text.get(1.0, tk.END).strip()
        self.msg_text.config(state=tk.DISABLED)
        
        full_text = f"ERROR MESSAGE:\n{message}\n"
        
        if self.tech_text:
            self.tech_text.config(state=tk.NORMAL)
            technical = self.tech_text.get(1.0, tk.END).strip()
            self.tech_text.config(state=tk.DISABLED)
            full_text += f"\nTECHNICAL DETAILS:\n{technical}"
        
        self.dialog.clipboard_clear()
        self.dialog.clipboard_append(full_text)
        self.dialog.update()
        
        self.show_copy_confirmation("All error information copied to clipboard!")
    
    def show_copy_confirmation(self, message):
        """Show a brief confirmation message with modern styling."""
        # Create a temporary label that fades out
        temp_frame = ttk.Frame(self.dialog, relief="solid", borderwidth=1)
        temp_frame.place(relx=0.5, rely=0.95, anchor=tk.CENTER)
        
        temp_label = ttk.Label(
            temp_frame, 
            text=f"{message}", 
            bootstyle="success",
            font=("Segoe UI", 10)
        )
        temp_label.pack(padx=15, pady=8)
        
        # Remove the label after 3 seconds
        self.dialog.after(3000, lambda: temp_frame.destroy())
    
    def close(self):
        """Close the dialog."""
        self.dialog.destroy()
    
    def show(self):
        """Show the dialog and wait for it to close."""
        self.dialog.wait_window()


class BoMApp:
    def __init__(self, root):
        self.root = root
        self.root.title("BoMination - BoM Processing Pipeline")
        self.root.geometry("700x800")  # Increased height for log panel
        self.root.resizable(True, True)

        self.pdf_path = tk.StringVar()
        self.page_range = tk.StringVar()
        self.company_name = tk.StringVar(value="")  # default to blank
        self.output_directory = tk.StringVar()  # Output directory selection

        # Progress and logging components
        self.progress_var = tk.StringVar(value="Ready to process your BoM files")
        self.progress_bar = None
        self.log_text = None
        
        self.build_gui()

    def build_gui(self):
        # Main container with padding
        main_container = ttk.Frame(self.root)
        main_container.pack(fill=BOTH, expand=True, padx=20, pady=20)
        
        # Title with modern styling
        title_label = ttk.Label(
            main_container, 
            text="BoMination", 
            font=("Segoe UI", 24, "bold")
        )
        title_label.pack(pady=(0, 5))
        
        subtitle_label = ttk.Label(
            main_container, 
            text="BoM Processing Pipeline", 
            font=("Segoe UI", 12),
            bootstyle="secondary"
        )
        subtitle_label.pack(pady=(0, 30))
        
        # Step 1: PDF File Selection
        pdf_frame = ttk.LabelFrame(main_container, text="Step 1: Select BoM PDF File", padding=15)
        pdf_frame.pack(fill=X, pady=(0, 15))
        
        pdf_entry_frame = ttk.Frame(pdf_frame)
        pdf_entry_frame.pack(fill=X, pady=5)
        
        ttk.Entry(
            pdf_entry_frame, 
            textvariable=self.pdf_path, 
            font=("Segoe UI", 10),
            width=50
        ).pack(side=LEFT, fill=X, expand=True, padx=(0, 10))
        
        ttk.Button(
            pdf_entry_frame, 
            text="Browse", 
            command=self.browse_pdf,
            bootstyle="outline-primary",
            width=10
        ).pack(side=RIGHT)
        
        # Step 2: Page Range
        page_frame = ttk.LabelFrame(main_container, text="Step 2: Enter Page Range", padding=15)
        page_frame.pack(fill=X, pady=(0, 15))
        
        page_entry_frame = ttk.Frame(page_frame)
        page_entry_frame.pack(fill=X, pady=5)
        
        ttk.Entry(
            page_entry_frame, 
            textvariable=self.page_range, 
            font=("Segoe UI", 10),
            width=20
        ).pack(side=LEFT, padx=(0, 10))
        
        # Help button with modern styling
        ttk.Button(
            page_entry_frame, 
            text="Help", 
            command=self.show_page_range_help,
            bootstyle="outline-info",
            width=8
        ).pack(side=RIGHT)
        
        # Examples label with better styling
        examples_label = ttk.Label(
            page_frame, 
            text="Examples: 1-3 (pages 1 to 3), 5 (page 5), 2,4,6 (pages 2, 4, and 6)", 
            font=("Segoe UI", 9),
            bootstyle="secondary"
        )
        examples_label.pack(anchor=W, pady=(5, 0))

        # Step 3: Company Selection
        company_frame = ttk.LabelFrame(main_container, text="Step 3: Select Company (Optional)", padding=15)
        company_frame.pack(fill=X, pady=(0, 15))
        
        company_dropdown = ttk.Combobox(
            company_frame,
            textvariable=self.company_name,
            values=["", "Farrell", "NEL"],
            state="readonly",
            font=("Segoe UI", 10),
            width=30
        )
        company_dropdown.pack(anchor=W, pady=5)
        company_dropdown.current(0)  # default to blank
        
        # Info label for company
        company_info_label = ttk.Label(
            company_frame, 
            text="Select if your PDF requires company-specific formatting", 
            font=("Segoe UI", 9),
            bootstyle="secondary"
        )
        company_info_label.pack(anchor=W, pady=(5, 0))

        # Step 4: Output Directory Selection
        output_frame = ttk.LabelFrame(main_container, text="Step 4: Choose Output Directory (Optional)", padding=15)
        output_frame.pack(fill=X, pady=(0, 20))
        
        output_entry_frame = ttk.Frame(output_frame)
        output_entry_frame.pack(fill=X, pady=5)
        
        ttk.Entry(
            output_entry_frame, 
            textvariable=self.output_directory, 
            font=("Segoe UI", 10),
            width=50
        ).pack(side=LEFT, fill=X, expand=True, padx=(0, 10))
        
        ttk.Button(
            output_entry_frame, 
            text="Browse", 
            command=self.browse_output_directory,
            bootstyle="outline-primary",
            width=10
        ).pack(side=RIGHT)
        
        # Info label for output directory
        output_info_label = ttk.Label(
            output_frame, 
            text="If not specified, files will be saved next to the input PDF", 
            font=("Segoe UI", 9),
            bootstyle="secondary"
        )
        output_info_label.pack(anchor=W, pady=(5, 0))

        # Action buttons frame
        button_frame = ttk.Frame(main_container)
        button_frame.pack(fill=X, pady=20)
        
        # Run Button with modern styling
        run_button = ttk.Button(
            button_frame, 
            text="Run Automation", 
            command=self.run_pipeline,
            bootstyle="success",
            width=20
        )
        run_button.pack(side=LEFT, padx=(0, 10))
        
        # Test button with updated styling
        test_button = ttk.Button(
            button_frame, 
            text="Test Error Dialog", 
            command=self.test_error_dialog,
            bootstyle="warning-outline",
            width=18
        )
        test_button.pack(side=LEFT)
        
        # Status/Progress area with modern card styling
        status_frame = ttk.Frame(main_container)
        status_frame.pack(fill=X, pady=(20, 0))
        
        # Progress card
        progress_card = ttk.LabelFrame(status_frame, text="Status & Progress", padding=15)
        progress_card.pack(fill=X, pady=(0, 10))
        
        # Status text
        self.status_label = ttk.Label(
            progress_card,
            textvariable=self.progress_var,
            font=("Segoe UI", 10),
            bootstyle="info"
        )
        self.status_label.pack(anchor=W, pady=(0, 10))
        
        # Progress bar (initially hidden)
        self.progress_bar = ttk.Progressbar(
            progress_card,
            mode="indeterminate",
            bootstyle="primary",
            length=400
        )
        self.progress_bar.pack(fill=X, pady=(0, 5))
        
        # Operation Log Panel
        log_frame = ttk.LabelFrame(main_container, text="Operation Log", padding=10)
        log_frame.pack(fill=BOTH, expand=True, pady=(10, 0))
        
        # Create scrollable text widget for logs
        log_scroll_frame = ttk.Frame(log_frame)
        log_scroll_frame.pack(fill=BOTH, expand=True)
        
        self.log_text = tk.Text(
            log_scroll_frame,
            wrap=tk.WORD,
            height=8,
            font=("Consolas", 9),
            relief="flat",
            borderwidth=0,
            bg="#f8f9fa",  # Light background
            state=tk.DISABLED
        )
        
        log_scrollbar = ttk.Scrollbar(log_scroll_frame, orient=VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scrollbar.set)
        
        self.log_text.pack(side=LEFT, fill=BOTH, expand=True)
        log_scrollbar.pack(side=RIGHT, fill=Y)
        
        # Log controls
        log_controls = ttk.Frame(log_frame)
        log_controls.pack(fill=X, pady=(10, 0))
        
        ttk.Button(
            log_controls,
            text="Copy Log",
            command=self.copy_log,
            bootstyle="outline-secondary",
            width=12
        ).pack(side=LEFT, padx=(0, 10))
        
        ttk.Button(
            log_controls,
            text="Clear Log",
            command=self.clear_log,
            bootstyle="outline-danger",
            width=12
        ).pack(side=LEFT)
        
        # Add initial log message
        self.add_log_message("Welcome to BoMination! Select a PDF file and configure settings to begin.", "info")

    def start_progress(self, message="Processing..."):
        """Start the progress bar with indeterminate mode."""
        def _start():
            self.progress_var.set(f"🔄 {message}")
            self.progress_bar.start(10)  # Update every 10ms for smooth animation
        
        # Ensure GUI update happens on main thread
        self.root.after(0, _start)
    
    def stop_progress(self, message="Ready"):
        """Stop the progress bar and update status."""
        def _stop():
            self.progress_bar.stop()
            self.progress_var.set(f"{message}")
        
        # Ensure GUI update happens on main thread
         
    
    def complete_progress(self):
        """Complete the progress and reset to ready state."""
        def _complete():
            self.progress_bar.stop()
            self.progress_var.set("✅ Ready for next operation")
        
        self.root.after(0, _complete)
    
    def update_status(self, message):
        """Update the status message without affecting progress bar."""
        def _update():
            self.progress_var.set(message)
        
        self.root.after(0, _update)
    
    def add_log_message(self, message, level="info"):
        """Add a timestamped message to the log panel."""
        def _add_log():
            timestamp = datetime.now().strftime("%H:%M:%S")
            
            # Color coding based on level
            level_icons = {
                "info": "ℹ️",
                "success": "✅", 
                "warning": "⚠️",
                "error": "❌",
                "step": "🔸"
            }
            
            icon = level_icons.get(level, "•")
            formatted_message = f"[{timestamp}] {icon} {message}\n"
            
            # Add to log text widget
            self.log_text.config(state=tk.NORMAL)
            self.log_text.insert(tk.END, formatted_message)
            self.log_text.see(tk.END)  # Auto-scroll to bottom
            self.log_text.config(state=tk.DISABLED)
        
        # Ensure GUI update happens on main thread
        self.root.after(0, _add_log)
    
    def copy_log(self):
        """Copy the entire log to clipboard."""
        self.log_text.config(state=tk.NORMAL)
        log_content = self.log_text.get(1.0, tk.END).strip()
        self.log_text.config(state=tk.DISABLED)
        
        if log_content:
            self.root.clipboard_clear()
            self.root.clipboard_append(log_content)
            self.root.update()
            self.add_log_message("Log copied to clipboard", "success")
        else:
            self.add_log_message("No log content to copy", "warning")
    
    def clear_log(self):
        """Clear the log panel."""
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state=tk.DISABLED)
        self.add_log_message("Log cleared", "info")

    def show_page_range_help(self):
        """Show detailed help for page range format."""
        help_text = """
📖 Page Range Format Help

Valid formats:
• Single page: 5
• Page range: 1-3 (pages 1, 2, and 3)
• Multiple pages: 2,4,6 (pages 2, 4, and 6)
• Mixed format: 1-3,5,7-9 (pages 1, 2, 3, 5, 7, 8, and 9)

Rules:
• Use only numbers, commas, and hyphens
• Page numbers must be positive integers
• In ranges, start page must be ≤ end page
• Spaces are allowed around commas and hyphens

Examples:
• "1" → Extract page 1 only
• "1-5" → Extract pages 1 through 5
• "2,4,7" → Extract pages 2, 4, and 7
• "1-3, 6, 8-10" → Extract pages 1, 2, 3, 6, 8, 9, and 10

Invalid formats:
• Empty or blank
• Letters: "a-b" 
• Negative numbers: "-1"
• Invalid ranges: "5-3"
        """
        messagebox.showinfo("Page Range Help", help_text.strip())

    def browse_pdf(self):
        file_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if file_path:
            self.pdf_path.set(file_path)
            self.add_log_message(f"Selected PDF: {Path(file_path).name}", "info")

    def browse_output_directory(self):
        """Browse for output directory selection."""
        directory = filedialog.askdirectory(title="Select Output Directory")
        if directory:
            self.output_directory.set(directory)
            self.add_log_message(f"Selected output directory: {directory}", "info")

    def run_pipeline(self):
        """Run the pipeline with comprehensive input validation."""
        pdf = self.pdf_path.get()
        pages = self.page_range.get()
        company = self.company_name.get()
        output_dir = self.output_directory.get()

        # Step 1: Validate PDF file
        pdf_valid, pdf_error = validate_pdf_file(pdf)
        if not pdf_valid:
            Messagebox.show_error("Invalid PDF File", pdf_error, parent=self.root)
            return

        # Step 2: Validate page range
        pages_valid, pages_error, parsed_ranges = validate_page_range(pages)
        if not pages_valid:
            Messagebox.show_error("Invalid Page Range", pages_error, parent=self.root)
            return

        # Step 3: Validate output directory
        output_valid, output_error = validate_output_directory(output_dir)
        if not output_valid:
            Messagebox.show_error("Invalid Output Directory", output_error, parent=self.root)
            return

        # Step 4: Check system requirements
        self.check_system_requirements()

        # Add template validation logging
        if COST_SHEET_TEMPLATE.exists():
            self.add_log_message(f"Cost sheet template found: {COST_SHEET_TEMPLATE.name}", "success")
        else:
            self.add_log_message(f"Cost sheet template not found: {COST_SHEET_TEMPLATE}", "warning")

        # If all validations pass, proceed with pipeline
        os.environ["BOM_PDF_PATH"] = pdf
        os.environ["BOM_PAGE_RANGE"] = pages
        os.environ["BOM_COMPANY"] = company
        os.environ["BOM_OUTPUT_DIRECTORY"] = self.output_directory.get() or ""

        # Log the start of pipeline
        self.add_log_message("Starting BoM processing pipeline...", "step")
        self.add_log_message(f"PDF: {Path(pdf).name}", "info")
        self.add_log_message(f"Pages: {pages}", "info")
        if company:
            self.add_log_message(f"Company: {company}", "info")
        if output_dir:
            self.add_log_message(f"Output directory: {output_dir}", "info")

        def background_task():
            try:
                # Start progress indication
                self.start_progress("Initializing pipeline...")
                
                self.add_log_message("Launching main pipeline process...", "step")
                
                # Update progress for different stages
                self.start_progress("Processing PDF and extracting tables...")
                
                # Call the pipeline with GUI review callback
                from main_pipeline import run_main_pipeline_with_gui_review
                
                # Create a callback that will be called when review is needed
                def review_callback(merged_df):
                    print("📝 GUI REVIEW: Review callback called from pipeline")
                    # This will be called from the background thread, so we need to
                    # schedule the review window creation on the main thread
                    result_container = [None]
                    review_done = [False]
                    
                    def show_review_on_main_thread():
                        try:
                            reviewed_df = self.show_review_window(merged_df)
                            result_container[0] = reviewed_df
                        except Exception as e:
                            print(f"📝 GUI REVIEW: Error in review window: {e}")
                            result_container[0] = merged_df  # Use original on error
                        finally:
                            review_done[0] = True
                    
                    # Schedule the review window on the main thread
                    self.root.after(0, show_review_on_main_thread)
                    
                    # Wait for the review to complete
                    while not review_done[0]:
                        time.sleep(0.1)
                    
                    return result_container[0]
                
                # Run the pipeline with the review callback
                result = run_main_pipeline_with_gui_review(
                    pdf_path=self.pdf_path.get(),
                    pages=pages,
                    company=company,
                    output_directory=self.output_directory.get(),
                    review_callback=review_callback
                )
                
                # Log the result if it contains useful information
                if result:
                    self.add_log_message(f"Pipeline result: {result}", "info")
                
                self.add_log_message("Pipeline process completed successfully!", "success")
                
                # Update progress to completion
                self.start_progress("Pipeline completed successfully!")
                self.complete_progress()
                
                # Stop progress and show success
                self.stop_progress("Pipeline completed successfully!")
                self.add_log_message("BoM processing pipeline completed successfully!", "success")
                
                # Look for output files in the PDF directory (changed for debugging)
                pdf_dir = Path(pdf).parent
                pdf_name = Path(pdf).stem
                
                # Expected output files
                expected_files = [
                    f"{pdf_name}_extracted.xlsx",
                    f"{pdf_name}_merged.xlsx", 
                    f"{pdf_name}_with_prices.xlsx",
                    f"{pdf_name}_cost_sheet.xlsx"
                ]
                
                found_files = []
                for filename in expected_files:
                    file_path = pdf_dir / filename
                    if file_path.exists():
                        found_files.append(str(file_path))
                        self.add_log_message(f"Output file: {filename}", "success")
                
                if found_files:
                    files_text = "\n".join([f"• {Path(f).name}" for f in found_files])
                    success_message = f"Pipeline completed successfully!\n\nOutput files created:\n{files_text}\n\nLocation: {pdf_dir}"
                else:
                    success_message = "Pipeline completed successfully!\n\nCheck the output folder for your processed files."
                
                # Schedule the success dialog on the main thread
                def show_success():
                    Messagebox.show_info(
                        "Success", 
                        success_message,
                        parent=self.root
                    )
                
                self.root.after(0, show_success)
                
            except Exception as e:
                # Stop progress on error
                self.stop_progress("Pipeline failed")
                self.add_log_message(f"Pipeline failed: {str(e)}", "error")
                
                print("=== Pipeline failed ===")
                print("Error:", str(e))
                
                # Schedule the error dialog on the main thread
                def show_error():
                    error_str = str(e).lower()
                    if "chromedriver" in error_str or "chrome" in error_str or "browser" in error_str:
                        # ChromeDriver specific error
                        CopyableErrorDialog(
                            self.root,
                            "ChromeDriver Error",
                            "The price lookup failed due to a ChromeDriver issue.\n\n" +
                            "This usually means:\n" +
                            "• ChromeDriver is not installed or not in the correct location\n" +
                            "• ChromeDriver version doesn't match your Chrome browser version\n" +
                            "• Chrome browser is not installed\n" +
                            "• Antivirus software is blocking ChromeDriver\n\n" +
                            "The pipeline completed the PDF extraction and table merging steps successfully, " +
                            "but could not retrieve pricing data from the web.\n\n" +
                            "To resolve this:\n" +
                            "1. Ensure Chrome browser is installed and up to date\n" +
                            "2. Download the matching ChromeDriver from: https://chromedriver.chromium.org/\n" +
                            "3. Place chromedriver.exe in the application's src folder\n" +
                            "4. Check that antivirus software isn't blocking the application",
                            f"Full error details:\n{str(e)}"
                        ).show()
                    else:
                        # Generic error dialog
                        Messagebox.show_error(
                            "Pipeline Error", 
                            f"Pipeline failed with error:\n\n{str(e)}",
                            parent=self.root
                        )
                
                self.root.after(0, show_error)

        threading.Thread(target=background_task).start()

    def check_system_requirements(self):
        """Check and warn about system requirements."""
        warnings = []
        
        # Check Java
        java_installed, java_version, java_error = check_java_installation()
        if not java_installed:
            self.add_log_message("Java not detected - required for PDF extraction", "warning")
            response = Messagebox.show_question(
                "Java Not Found",
                "Java is required for PDF table extraction.\n\n" + 
                (java_error or "Java not detected.") + 
                "\n\nWould you like to download Java now?",
                parent=self.root
            )
            if response == "Yes":
                open_help_url("https://www.java.com/download/")
            warnings.append("Java not installed")
        else:
            self.add_log_message(f"Java detected: {java_version}", "success")
        
        # Check ChromeDriver
        chrome_available, chrome_version, chrome_error = check_chromedriver_availability()
        if not chrome_available:
            self.add_log_message("ChromeDriver not available - required for price lookup", "warning")
            self.add_log_message(f"ChromeDriver error: {chrome_error}", "error")
            response = Messagebox.show_question(
                "ChromeDriver Not Found",
                "ChromeDriver is required for price lookup.\n\n" + 
                (chrome_error or "ChromeDriver not detected.") + 
                "\n\nWould you like to open the ChromeDriver download page?",
                parent=self.root
            )
            if response == "Yes":
                open_help_url("https://chromedriver.chromium.org/downloads")
            warnings.append("ChromeDriver not available")
        else:
            self.add_log_message(f"ChromeDriver detected: {chrome_version}", "success")
        
        # Show summary if there are warnings
        if warnings:
            warning_text = "System Requirements Warning:\n\n" + "\n".join([f"• {w}" for w in warnings])
            warning_text += "\n\nYou can still try to run the pipeline, but some features may not work properly."
            Messagebox.show_warning("System Requirements", warning_text, parent=self.root)
        else:
            self.add_log_message("All system requirements met", "success")

    def test_error_dialog(self):
        """Test the copyable error dialog (for development/testing)."""
        self.add_log_message("Testing error dialog functionality", "info")
        
        sample_error = """Sample Error for Testing

This is a sample error message to test the copyable error dialog functionality.

Possible solutions:
• Solution 1: Try this first
• Solution 2: If that doesn't work, try this
• Solution 3: As a last resort, try this

This error dialog allows you to copy the error message to share with support or for troubleshooting."""

        sample_technical = """Command: ['python', 'main_pipeline.py']
Return code: 1

Output:
Sample output from the failed command
Multiple lines of output
With various information

Errors:
Sample error output
Error details
Stack trace information"""

        try:
            error_dialog = CopyableErrorDialog(
                self.root,
                "Test Error Dialog",
                sample_error,
                sample_technical
            )
            error_dialog.show()
            self.add_log_message("Error dialog test completed", "success")
        except Exception as e:
            self.add_log_message(f"Error dialog test failed: {e}", "error")
            messagebox.showerror("Error", f"Failed to show error dialog: {e}")

    def show_review_window(self, merged_df):
        """
        Show the review window for the merged BoM table.
        This runs in the main GUI thread.
        """
        print(f"📝 GUI REVIEW: Creating review window for dataframe with shape {merged_df.shape}")
        
        # Create a new top-level window
        review_window = tk.Toplevel(self.root)
        review_window.title("Review and Edit Merged BoM Table")
        
        # Make the window fullscreen
        review_window.attributes('-fullscreen', True)
        review_window.configure(bg='white')
        
        # Add escape key to exit fullscreen
        review_window.bind('<Escape>', lambda e: review_window.attributes('-fullscreen', False))
        review_window.bind('<F11>', lambda e: review_window.attributes('-fullscreen', not review_window.attributes('-fullscreen')))
        
        # Make window modal and focused
        review_window.transient(self.root)
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
                'multipleselectioncolor': '#CCE7FF'
            }
            
            # Create table with custom options
            table_widget = Table(table_frame, dataframe=merged_df, **table_options)
            table_widget.show()
            
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

if __name__ == "__main__":
    try:
        # Create the main window with ttkbootstrap's darkly theme
        root = ttk.Window(themename="darkly")
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
    finally:
        # Force cleanup
        try:
            if 'root' in locals():
                root.quit()
                root.destroy()
        except:
            pass