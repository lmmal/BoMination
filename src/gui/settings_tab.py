import tkinter as tk
from tkinter import filedialog
import ttkbootstrap as ttk
from ttkbootstrap.constants import *


class SettingsTab:
    """Settings tab for BoMination application configuration."""
    
    def __init__(self, parent_frame, app_instance):
        self.parent_frame = parent_frame
        self.app = app_instance
        self.build_settings_gui()
    
    def build_settings_gui(self):
        """Build the settings tab interface."""
        # Main container with padding
        main_container = ttk.Frame(self.parent_frame)
        main_container.pack(fill=BOTH, expand=True, padx=20, pady=20)
        
        # Settings title
        title_label = ttk.Label(
            main_container, 
            text="Settings & Configuration", 
            font=("Segoe UI", 18, "bold")
        )
        title_label.pack(pady=(0, 20))
        
        # Table Detection Mode Section
        detection_frame = ttk.LabelFrame(main_container, text="Table Detection Mode", padding=15)
        detection_frame.pack(fill=X, pady=(0, 15))
        
        # Detection mode dropdown
        detection_dropdown = ttk.Combobox(
            detection_frame,
            textvariable=self.app.tabula_mode,
            values=["conservative", "balanced", "aggressive"],
            state="readonly",
            font=("Segoe UI", 10),
            width=30
        )
        detection_dropdown.pack(anchor=W, pady=5)
        detection_dropdown.current(1)  # default to balanced
        
        # Info label for tabula mode
        detection_info_label = ttk.Label(
            detection_frame, 
            text="• Conservative: Fewer false positives, may miss some tables\n"
                 "• Balanced: Good compromise between accuracy and completeness (Recommended)\n"
                 "• Aggressive: Detects more tables, but may include non-table content", 
            font=("Segoe UI", 9),
            bootstyle="secondary",
            justify=tk.LEFT
        )
        detection_info_label.pack(anchor=W, pady=(5, 0))
        
        # Manual Table Area Selection Section
        roi_frame = ttk.LabelFrame(main_container, text="Manual Table Area Selection", padding=15)
        roi_frame.pack(fill=X, pady=(0, 15))
        
        # ROI checkbox
        roi_checkbox = ttk.Checkbutton(
            roi_frame,
            text="Enable manual table area selection (ROI)",
            variable=self.app.use_roi,
            bootstyle="primary"
        )
        roi_checkbox.pack(anchor=W, pady=5)
        
        # Info label for ROI mode
        roi_info_label = ttk.Label(
            roi_frame, 
            text="When enabled, you can manually select table areas on each page for more precise extraction.\n"
                 "This is useful when automatic table detection fails or you need to select specific regions.\n"
                 "💡 TIP: Try ROI mode if automatic extraction finds tables but they contain poor quality text.",
            font=("Segoe UI", 9),
            bootstyle="secondary",
            justify=tk.LEFT
        )
        roi_info_label.pack(anchor=W, pady=(5, 0))
        
        # Output Directory Section
        output_frame = ttk.LabelFrame(main_container, text="Output Directory", padding=15)
        output_frame.pack(fill=X, pady=(0, 15))
        
        # Output directory entry and browse button
        output_entry_frame = ttk.Frame(output_frame)
        output_entry_frame.pack(fill=X, pady=5)
        
        ttk.Entry(
            output_entry_frame, 
            textvariable=self.app.output_directory, 
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
            text="If not specified, files will be saved next to the input PDF file.\n"
                 "Choose a directory to save all output files to a specific location.", 
            font=("Segoe UI", 9),
            bootstyle="secondary",
            justify=tk.LEFT
        )
        output_info_label.pack(anchor=W, pady=(5, 0))
        
        # Advanced Settings Section
        advanced_frame = ttk.LabelFrame(main_container, text="Advanced Settings", padding=15)
        advanced_frame.pack(fill=X, pady=(0, 15))
        
        # System requirements check button
        system_check_button = ttk.Button(
            advanced_frame,
            text="Check System Requirements",
            command=self.check_system_requirements,
            bootstyle="outline-info",
            width=25
        )
        system_check_button.pack(anchor=W, pady=5)
        
        # Reset to defaults button
        reset_button = ttk.Button(
            advanced_frame,
            text="Reset to Defaults",
            command=self.reset_to_defaults,
            bootstyle="outline-warning",
            width=20
        )
        reset_button.pack(anchor=W, pady=(5, 0))
        
        # Info about advanced settings
        advanced_info_label = ttk.Label(
            advanced_frame,
            text="• System Requirements: Check for Java, ChromeDriver, and OCR tools\n"
                 "• Reset to Defaults: Restore all settings to their original values",
            font=("Segoe UI", 9),
            bootstyle="secondary",
            justify=tk.LEFT
        )
        advanced_info_label.pack(anchor=W, pady=(5, 0))
    
    def browse_output_directory(self):
        """Browse for output directory selection."""
        directory = filedialog.askdirectory(title="Select Output Directory")
        if directory:
            self.app.output_directory.set(directory)
            self.app.add_log_message(f"Selected output directory: {directory}", "info")
    
    def check_system_requirements(self):
        """Run system requirements check with popups enabled."""
        self.app.add_log_message("Checking system requirements...", "info")
        self.app.check_system_requirements(silent=False)
    
    def reset_to_defaults(self):
        """Reset all settings to their default values."""
        self.app.tabula_mode.set("balanced")
        self.app.use_roi.set(False)
        self.app.output_directory.set("")
        self.app.add_log_message("Settings reset to defaults", "info")