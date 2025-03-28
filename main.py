#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Moxy Rates Template Transfer Application

This application automates the process of transferring data from an Adjusted Rates
spreadsheet into a Template spreadsheet, with intelligent column mapping and format detection.
"""

import os
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import configparser
import threading
import queue
import logging
from datetime import datetime
import pandas as pd
import subprocess
from PIL import Image, ImageTk
import re

# Try to import xlsxwriter - used as fallback for Excel saving
try:
    import xlsxwriter
    XLSXWRITER_AVAILABLE = True
except ImportError:
    XLSXWRITER_AVAILABLE = False
    logging.warning("xlsxwriter package not found. Will use default Excel engines only.")

# Import custom modules
from file_analyzer import FileAnalyzer
from mapping_system import MappingSystem, MappingDialog
from config_manager import ConfigManager, MappingConfigManager
from data_processor import DataProcessor

class Application(tk.Tk):
    """Main application window for Moxy Rates Template Transfer."""
    
    # Color constants for Moxy theme
    DARK_BG = "#1E1E1E"  # Darker background
    DARKER_BG = "#171717"  # Even darker for contrast
    BUTTON_BG = "#1B4B8F"  # Moxy blue
    BUTTON_ACTIVE_BG = "#8DC63F"  # Moxy green
    BUTTON_HOVER_BG = "#2B5BA0"  # Slightly lighter blue for hover
    BUTTON_DISABLED_BG = "#2A2A2A"  # Subtle disabled state
    ENTRY_BG = "#2D2D2D"  # Slightly lighter than background for input fields
    TEXT_COLOR = "#E0E0E0"  # Bright text for contrast
    TEXT_DISABLED_COLOR = "#666666"  # Dimmed text
    ACCENT_COLOR = "#8DC63F"  # Moxy green for accents
    BORDER_COLOR = "#333333"  # Subtle borders
    PROGRESS_BG = "#8DC63F"  # Moxy green for progress
    PROGRESS_TROUGH = "#1B4B8F"  # Moxy blue for progress background
    
    def __init__(self):
        """Initialize the application."""
        super().__init__()
        
        # Set window properties for modern look
        self.title("Moxy Rates Template Transfer")
        self.geometry("900x700")  # Slightly taller for better spacing
        self.minsize(800, 600)
        
        # Configure the window background
        self.configure(background=self.DARK_BG)
        
        # Set theme to clam for better styling control
        style = ttk.Style()
        if style.theme_use() != 'clam':
            style.theme_use('clam')
        
        # Configure system-level styles for tk widgets
        self.option_add('*TCombobox*Listbox.background', self.DARKER_BG)
        self.option_add('*TCombobox*Listbox.foreground', '#FFFFFF')
        self.option_add('*TCombobox*Listbox.selectBackground', self.BUTTON_BG)
        self.option_add('*TCombobox*Listbox.selectForeground', '#FFFFFF')
        self.option_add('*TCombobox*background', self.DARKER_BG)
        self.option_add('*TCombobox*foreground', '#FFFFFF')
        self.option_add('*TCombobox*fieldBackground', self.DARKER_BG)
        self.option_add('*TCombobox*selectBackground', self.BUTTON_BG)
        self.option_add('*TCombobox*selectForeground', '#FFFFFF')
        self.option_add('*TCombobox*arrowColor', '#FFFFFF')
        self.option_add('*Entry.background', self.DARKER_BG)
        self.option_add('*Entry.foreground', '#FFFFFF')
        self.option_add('*Entry.insertBackground', '#FFFFFF')
        self.option_add('*Listbox.background', self.DARKER_BG)
        self.option_add('*Listbox.foreground', '#FFFFFF')
        self.option_add('*Listbox.selectBackground', self.BUTTON_BG)
        self.option_add('*Listbox.selectForeground', '#FFFFFF')
        
        # Initialize configuration and components
        self.config_mgr = ConfigManager()
        self.config_mgr.load_config()
        self.configure_logging()
        self.file_analyzer = FileAnalyzer()
        self.data_processor = DataProcessor()
        self.mapping_system = MappingSystem(self.config_mgr)
        self.mapping_config = MappingConfigManager()
        
        # Set up custom styles
        self.setup_styles()
        
        # Create main container with padding
        main_frame = ttk.Frame(self, style="Padded.TFrame")
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Create UI sections with spacing
        self.create_input_section(main_frame)
        ttk.Frame(main_frame, height=10, style="TFrame").pack()  # Spacer
        self.create_output_section(main_frame)
        ttk.Frame(main_frame, height=10, style="TFrame").pack()  # Spacer
        self.create_options_section(main_frame)
        ttk.Frame(main_frame, height=10, style="TFrame").pack()  # Spacer
        self.create_button_section(main_frame)
        self.create_status_section(main_frame)
        
        # Message queue for thread-safe UI updates
        self.msg_queue = queue.Queue()
        self.after(100, self.process_queue)
        
        # Load saved settings
        self.load_settings()
        
        # Update status
        self.status_var.set("Ready")
        logging.info("Application initialized")
    
    def configure_logging(self):
        """Configure the logging system."""
        # Create logs directory if it doesn't exist
        script_dir = os.path.dirname(os.path.abspath(__file__))
        log_dir = os.path.join(script_dir, "logs")
        if not os.path.exists(log_dir):
            os.makedirs(log_dir)
            
        log_file = os.path.join(log_dir, f"app_{datetime.now().strftime('%Y%m%d')}.log")
        
        # Force create a new log file to check write permissions
        try:
            with open(log_file, 'w') as f:
                f.write(f"Log file created at {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            print(f"Successfully created log file: {log_file}")
        except Exception as e:
            print(f"WARNING: Could not create log file {log_file}: {str(e)}")
            # Try to create in current directory as fallback
            log_file = f"app_{datetime.now().strftime('%Y%m%d')}.log"
            
        # Configure logging
        log_level = logging.DEBUG if self.config_mgr.get_setting("enable_logging", False) else logging.INFO
        
        # Clear any existing handlers
        root = logging.getLogger()
        for handler in root.handlers[:]:
            root.removeHandler(handler)
            
        # Create new handlers
        file_handler = logging.FileHandler(log_file, mode='a')
        console_handler = logging.StreamHandler()
        
        # Create formatter
        formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
        file_handler.setFormatter(formatter)
        console_handler.setFormatter(formatter)
        
        # Add handlers
        root.addHandler(file_handler)
        root.addHandler(console_handler)
        
        # Set log level
        root.setLevel(log_level)
        
        # Test logging
        logging.info("Logging configured successfully")
        logging.debug("Debug logging is enabled")
    
    def create_ui(self):
        """Create the user interface."""
        # Create main frame
        main_frame = ttk.Frame(self, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Input Files section
        self.create_input_section(main_frame)
        
        # Output section
        self.create_output_section(main_frame)
        
        # Options section
        self.create_options_section(main_frame)
        
        # Buttons section
        self.create_button_section(main_frame)
        
        # Status section
        self.create_status_section(main_frame)
    
    def create_input_section(self, parent):
        """Create the input files section with modern styling."""
        input_frame = ttk.LabelFrame(parent, text="Input Files", padding="15", style="TLabelframe")
        input_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Create grid with proper spacing
        input_frame.columnconfigure(1, weight=1)  # Make the entry column expandable
        
        # Adjusted Rates file row
        ttk.Label(input_frame, text="Adjusted Rates File:", 
                 style="TLabel").grid(row=0, column=0, sticky=tk.W, pady=(5, 10))
        
        # Create a frame for the entry and button
        rates_frame = ttk.Frame(input_frame, style="TFrame")
        rates_frame.grid(row=0, column=1, columnspan=2, sticky=tk.EW, pady=(5, 10), padx=(5, 0))
        rates_frame.columnconfigure(0, weight=1)
        
        self.adjusted_rates_var = tk.StringVar()
        entry1 = tk.Entry(rates_frame, textvariable=self.adjusted_rates_var,
                         bg=self.DARKER_BG, fg='#FFFFFF',
                         insertbackground='#FFFFFF',
                         relief='solid', bd=1)
        entry1.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        
        browse_btn1 = tk.Button(rates_frame, text="Browse", command=self.browse_adjusted_rates,
                              bg=self.BUTTON_BG, fg='#FFFFFF',
                              font=("Segoe UI", 9),
                              relief='flat',
                              activebackground=self.BUTTON_HOVER_BG,
                              activeforeground='#FFFFFF',
                              width=8)
        browse_btn1.pack(side=tk.RIGHT)
        
        # Template file row
        ttk.Label(input_frame, text="Template File:", 
                 style="TLabel").grid(row=1, column=0, sticky=tk.W, pady=(0, 5))
        
        # Create a frame for the entry and button
        template_frame = ttk.Frame(input_frame, style="TFrame")
        template_frame.grid(row=1, column=1, columnspan=2, sticky=tk.EW, pady=(0, 5), padx=(5, 0))
        template_frame.columnconfigure(0, weight=1)
        
        self.template_var = tk.StringVar()
        entry2 = tk.Entry(template_frame, textvariable=self.template_var,
                         bg=self.DARKER_BG, fg='#FFFFFF',
                         insertbackground='#FFFFFF',
                         relief='solid', bd=1)
        entry2.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        
        browse_btn2 = tk.Button(template_frame, text="Browse", command=self.browse_template,
                              bg=self.BUTTON_BG, fg='#FFFFFF',
                              font=("Segoe UI", 9),
                              relief='flat',
                              activebackground=self.BUTTON_HOVER_BG,
                              activeforeground='#FFFFFF',
                              width=8)
        browse_btn2.pack(side=tk.RIGHT)
    
    def create_output_section(self, parent):
        """Create the output section with modern styling."""
        output_frame = ttk.LabelFrame(parent, text="Output", padding="15", style="TLabelframe")
        output_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Create grid with proper spacing
        output_frame.columnconfigure(1, weight=1)  # Make the entry column expandable
        
        # Output file row with modern layout
        ttk.Label(output_frame, text="Output Filename:", 
                 style="TLabel").grid(row=0, column=0, sticky=tk.W, pady=5)
        
        # Create a frame for the entry and button
        file_frame = ttk.Frame(output_frame, style="TFrame")
        file_frame.grid(row=0, column=1, columnspan=2, sticky=tk.EW, pady=5, padx=(5, 0))
        file_frame.columnconfigure(0, weight=1)
        
        self.output_var = tk.StringVar()
        entry = tk.Entry(file_frame, textvariable=self.output_var,
                        bg=self.DARKER_BG, fg='#FFFFFF',
                        insertbackground='#FFFFFF',
                        relief='solid', bd=1)
        entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        
        browse_btn = tk.Button(file_frame, text="Browse", command=self.browse_output,
                             bg=self.BUTTON_BG, fg='#FFFFFF',
                             font=("Segoe UI", 9),
                             relief='flat',
                             activebackground=self.BUTTON_HOVER_BG,
                             activeforeground='#FFFFFF',
                             width=8)
        browse_btn.pack(side=tk.RIGHT)
    
    def create_options_section(self, parent):
        """Create the options section with modern styling."""
        options_frame = ttk.LabelFrame(parent, text="Options", padding="15", style="TLabelframe")
        options_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Main options container
        main_options = ttk.Frame(options_frame, style="TFrame")
        main_options.pack(fill=tk.X, expand=True, pady=(0, 10))
        main_options.columnconfigure(0, weight=1)
        main_options.columnconfigure(1, weight=1)
        
        # Left column for checkboxes
        left_col = ttk.Frame(main_options, style="TFrame", padding=(0, 0, 10, 0))
        left_col.grid(row=0, column=0, sticky=tk.EW)
        
        # Remember directories option
        self.remember_dirs_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(left_col, text="Remember last used directories",
                       variable=self.remember_dirs_var,
                       style="TCheckbutton").pack(anchor=tk.W, pady=2)
        
        # Auto-open output option
        self.auto_open_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(left_col, text="Open output file after processing",
                       variable=self.auto_open_var,
                       style="TCheckbutton").pack(anchor=tk.W, pady=2)
        
        # Enable logging option
        self.logging_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(left_col, text="Enable detailed logging",
                       variable=self.logging_var,
                       style="TCheckbutton").pack(anchor=tk.W, pady=2)
        
        # Right column for checkboxes
        right_col = ttk.Frame(main_options, style="TFrame", padding=(10, 0, 0, 0))
        right_col.grid(row=0, column=1, sticky=tk.EW)
        
        # Auto-detect format option
        self.auto_detect_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(right_col, text="Auto-detect file format",
                       variable=self.auto_detect_var,
                       style="TCheckbutton").pack(anchor=tk.W, pady=2)
        
        # Use saved mappings option
        self.use_saved_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(right_col, text="Use saved mappings when available",
                       variable=self.use_saved_var,
                       style="TCheckbutton").pack(anchor=tk.W, pady=2)
        
        # Enhanced formatting option
        self.enhanced_format_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(right_col, text="Enhanced formatting",
                       variable=self.enhanced_format_var,
                       style="TCheckbutton").pack(anchor=tk.W, pady=2)
        
        # Separator line
        separator = ttk.Frame(options_frame, height=1, style="Separator.TFrame")
        separator.pack(fill=tk.X, pady=10)
        
        # Bottom frame for sheet selection and default deductible
        bottom_frame = ttk.Frame(options_frame, style="TFrame")
        bottom_frame.pack(fill=tk.X, expand=True)
        bottom_frame.columnconfigure(0, weight=1)
        bottom_frame.columnconfigure(1, weight=1)
        
        # Left side - Sheet selection
        sheet_frame = ttk.Frame(bottom_frame, style="TFrame")
        sheet_frame.grid(row=0, column=0, sticky=tk.EW, padx=(0, 10))
        
        ttk.Label(sheet_frame, text="Adjusted Rates Sheet:",
                 style="TLabel").pack(side=tk.TOP, anchor=tk.W, pady=(0, 2))
        self.adjusted_sheet_var = tk.StringVar()
        self.adjusted_sheet_combo = ttk.Combobox(sheet_frame,
                                               textvariable=self.adjusted_sheet_var,
                                               style="Dark.TCombobox",
                                               state="readonly")
        self.adjusted_sheet_combo.pack(side=tk.TOP, fill=tk.X, pady=(0, 5))
        
        ttk.Label(sheet_frame, text="Template Sheet:",
                 style="TLabel").pack(side=tk.TOP, anchor=tk.W, pady=(0, 2))
        self.template_sheet_var = tk.StringVar()
        self.template_sheet_combo = ttk.Combobox(sheet_frame,
                                              textvariable=self.template_sheet_var,
                                              style="Dark.TCombobox",
                                              state="readonly")
        self.template_sheet_combo.pack(side=tk.TOP, fill=tk.X)
        
        # Right side - Default Deductible
        deduct_frame = ttk.Frame(bottom_frame, style="TFrame")
        deduct_frame.grid(row=0, column=1, sticky=tk.EW, padx=(10, 0))
        
        # Create a header frame for the label and help icon
        deduct_header = ttk.Frame(deduct_frame, style="TFrame")
        deduct_header.pack(fill=tk.X, anchor=tk.W, pady=(0, 2))
        
        ttk.Label(deduct_header, text="Default Deductible:",
                 style="TLabel").pack(side=tk.LEFT)
        
        help_label = ttk.Label(deduct_header, text="ⓘ", style="TLabel",
                            cursor="question_arrow")
        help_label.pack(side=tk.LEFT, padx=(5, 0))
        
        # Add tooltip
        help_text = "This will be used as the preferred value for the PlanDeduct column.\nIf a row doesn't have this deductible, the lowest available will be used."
        self._create_tooltip(help_label, help_text)
        
        # Entry field for deductible
        self.default_deduct_var = tk.StringVar(value="100")
        deduct_entry = tk.Entry(deduct_frame, textvariable=self.default_deduct_var,
                              width=15, bg=self.DARKER_BG, fg='#FFFFFF',
                              insertbackground='#FFFFFF',
                              relief='solid', bd=1)
        deduct_entry.pack(fill=tk.X, pady=(0, 5))
    
    def create_button_section(self, parent):
        """Create the action buttons section with modern styling."""
        button_frame = ttk.Frame(parent, style="TFrame", padding="5")
        button_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Left side buttons
        left_buttons = ttk.Frame(button_frame, style="TFrame")
        left_buttons.pack(side=tk.LEFT)
        
        # Process button with Moxy green styling
        process_btn = tk.Button(left_buttons, text="Process Files", 
                              command=self.process_files,
                              bg=self.ACCENT_COLOR,
                              fg='#000000',
                              font=("Segoe UI", 9, "bold"),
                              relief='flat',
                              activebackground="#9ED84F",
                              activeforeground='#000000',
                              width=15)
        process_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        # Preview mapping button with Moxy blue styling
        preview_btn = tk.Button(left_buttons, text="Preview Mapping", 
                              command=self.preview_mapping,
                              bg=self.BUTTON_BG,
                              fg='#FFFFFF',
                              font=("Segoe UI", 9, "bold"),
                              relief='flat',
                              activebackground=self.BUTTON_HOVER_BG,
                              activeforeground='#FFFFFF',
                              width=15)
        preview_btn.pack(side=tk.LEFT, padx=5)
        
        # Right side buttons
        right_buttons = ttk.Frame(button_frame, style="TFrame")
        right_buttons.pack(side=tk.RIGHT)
        
        # Exit button with red accent
        exit_btn = tk.Button(right_buttons, text="Exit", 
                           command=self.on_exit,
                           bg=self.BUTTON_BG,
                           fg='#FFFFFF',
                           font=("Segoe UI", 9, "bold"),
                           relief='flat',
                           activebackground=self.BUTTON_HOVER_BG,
                           activeforeground='#FFFFFF',
                           width=10)
        exit_btn.pack(side=tk.RIGHT, padx=(5, 0))
    
    def create_status_section(self, parent):
        """Create the status section with modern styling."""
        # Create a styled frame with a subtle border
        status_frame = ttk.Frame(parent, style="Status.TFrame", padding="15")
        status_frame.pack(fill=tk.X, pady=(0, 10), side=tk.BOTTOM)
        
        # Status section with improved layout
        status_label_frame = ttk.Frame(status_frame, style="Status.TFrame")
        status_label_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Status label with better styling
        ttk.Label(status_label_frame, text="Status:", 
                 style="Subheader.TLabel").pack(side=tk.LEFT)
        
        self.status_var = tk.StringVar(value="Ready")
        status_display = ttk.Label(status_label_frame, 
                                 textvariable=self.status_var,
                                 style="Status.TLabel")
        status_display.pack(side=tk.LEFT, padx=10, fill=tk.X, expand=True)
        
        # Progress section with modern appearance
        progress_frame = ttk.Frame(status_frame, style="Status.TFrame")
        progress_frame.pack(fill=tk.X)
        progress_frame.columnconfigure(0, weight=1)  # Make progress bar expandable
        
        # Progress bar with increased height and Moxy colors
        self.progress_var = tk.DoubleVar(value=0.0)
        self.progress_percentage = tk.StringVar(value="0%")
        
        self.progress_bar = ttk.Progressbar(
            progress_frame, 
            variable=self.progress_var,
            mode="determinate",
            style="Horizontal.TProgressbar"
        )
        self.progress_bar.grid(row=0, column=0, sticky=tk.EW, padx=(0, 10))
        
        # Progress percentage with modern styling
        percentage_label = ttk.Label(progress_frame, 
                                   textvariable=self.progress_percentage,
                                   style="Status.TLabel",
                                   width=5)
        percentage_label.grid(row=0, column=1)
    
    def load_settings(self):
        """Load saved settings from config."""
        # Load checkbox settings
        self.remember_dirs_var.set(self.config_mgr.get_setting("remember_directories", True))
        self.auto_open_var.set(self.config_mgr.get_setting("open_after_processing", True))
        self.logging_var.set(self.config_mgr.get_setting("enable_logging", True))
        self.auto_detect_var.set(self.config_mgr.get_setting("auto_detect_formats", True))
        self.use_saved_var.set(self.config_mgr.get_setting("use_saved_mappings", True))
        self.enhanced_format_var.set(self.config_mgr.get_setting("enhanced_detection", True))
        
        # Load sheet names
        self.adjusted_sheet_var.set(self.config_mgr.get_setting("adjusted_sheet_name", "Dealer Cost Rates"))
        self.template_sheet_var.set(self.config_mgr.get_setting("template_sheet_name", "Sheet1"))
        
        # Load directory paths if remember is enabled
        if self.remember_dirs_var.get():
            last_adjusted_dir = self.config_mgr.get_setting("last_adjusted_rates_dir", "")
            last_template_dir = self.config_mgr.get_setting("last_template_dir", "")
            last_output_dir = self.config_mgr.get_setting("last_output_dir", "")
            
            self.adjusted_rates_var.set(self.config_mgr.get_setting("last_adjusted_rates_file", ""))
            self.template_var.set(self.config_mgr.get_setting("last_template_file", ""))
            self.output_var.set(self.config_mgr.get_setting("last_output_file", ""))
    
    def save_settings(self):
        """Save current settings to config."""
        # Save checkbox settings
        self.config_mgr.set_setting("remember_directories", self.remember_dirs_var.get())
        self.config_mgr.set_setting("open_after_processing", self.auto_open_var.get())
        self.config_mgr.set_setting("enable_logging", self.logging_var.get())
        self.config_mgr.set_setting("auto_detect_formats", self.auto_detect_var.get())
        self.config_mgr.set_setting("use_saved_mappings", self.use_saved_var.get())
        self.config_mgr.set_setting("enhanced_detection", self.enhanced_format_var.get())
        
        # Save sheet names
        self.config_mgr.set_setting("adjusted_sheet_name", self.adjusted_sheet_var.get())
        self.config_mgr.set_setting("template_sheet_name", self.template_sheet_var.get())
        
        # Save directory paths if remember is enabled
        if self.remember_dirs_var.get():
            adjusted_file = self.adjusted_rates_var.get()
            template_file = self.template_var.get()
            output_file = self.output_var.get()
            
            self.config_mgr.set_setting("last_adjusted_rates_file", adjusted_file)
            self.config_mgr.set_setting("last_template_file", template_file)
            self.config_mgr.set_setting("last_output_file", output_file)
            
            if adjusted_file:
                self.config_mgr.set_setting("last_adjusted_rates_dir", os.path.dirname(adjusted_file))
            if template_file:
                self.config_mgr.set_setting("last_template_dir", os.path.dirname(template_file))
            if output_file:
                self.config_mgr.set_setting("last_output_dir", os.path.dirname(output_file))
        
        # Save configuration
        self.config_mgr.save_config()
    
    def browse_adjusted_rates(self):
        """Browse for adjusted rates file."""
        last_dir = self.config_mgr.get_setting("last_adjusted_rates_dir", os.path.expanduser("~"))
        filename = filedialog.askopenfilename(
            title="Select Adjusted Rates File",
            initialdir=last_dir if os.path.exists(last_dir) else os.path.expanduser("~"),
            filetypes=[("Excel Files", "*.xlsx *.xls"), ("All Files", "*.*")]
        )
        
        if filename:
            self.adjusted_rates_var.set(filename)
            self.update_adjusted_sheets()
            
            # Generate default output filename
            if not self.output_var.get():
                base_name = os.path.splitext(os.path.basename(filename))[0]
                output_dir = self.config_mgr.get_setting("last_output_dir", os.path.dirname(filename))
                self.output_var.set(os.path.join(output_dir, f"{base_name}_processed.xlsx"))
    
    def browse_template(self):
        """Browse for template file."""
        last_dir = self.config_mgr.get_setting("last_template_dir", os.path.expanduser("~"))
        filename = filedialog.askopenfilename(
            title="Select Template File",
            initialdir=last_dir if os.path.exists(last_dir) else os.path.expanduser("~"),
            filetypes=[("Excel Files", "*.xlsx *.xls"), ("All Files", "*.*")]
        )
        
        if filename:
            self.template_var.set(filename)
            self.update_template_sheets()
    
    def browse_output(self):
        """Browse for output file location."""
        last_dir = self.config_mgr.get_setting("last_output_dir", os.path.expanduser("~"))
        filename = filedialog.asksaveasfilename(
            title="Save Output As",
            initialdir=last_dir if os.path.exists(last_dir) else os.path.expanduser("~"),
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx"), ("Excel 97-2003", "*.xls"), ("All Files", "*.*")]
        )
        
        if filename:
            # Ensure we have a valid Excel extension
            if not filename.lower().endswith(('.xlsx', '.xls')):
                filename += '.xlsx'
                logging.info(f"Added .xlsx extension to output filename: {filename}")
            
            self.output_var.set(filename)
    
    def update_adjusted_sheets(self):
        """Update the available sheets in adjusted rates file."""
        try:
            filename = self.adjusted_rates_var.get()
            if filename and os.path.exists(filename):
                sheets = self.file_analyzer.get_sheet_names(filename)
                self.adjusted_sheet_combo['values'] = sheets
                
                # Set to first sheet if current selection not in list
                if self.adjusted_sheet_var.get() not in sheets and sheets:
                    self.adjusted_sheet_var.set(sheets[0])
        except Exception as e:
            logging.error(f"Error updating adjusted sheets: {str(e)}")
            messagebox.showerror("Error", f"Error reading sheets: {str(e)}")
    
    def update_template_sheets(self):
        """Update the available sheets in template file."""
        try:
            filename = self.template_var.get()
            if filename and os.path.exists(filename):
                sheets = self.file_analyzer.get_sheet_names(filename)
                self.template_sheet_combo['values'] = sheets
                
                # Set to first sheet if current selection not in list
                if self.template_sheet_var.get() not in sheets and sheets:
                    self.template_sheet_var.set(sheets[0])
        except Exception as e:
            logging.error(f"Error updating template sheets: {str(e)}")
            messagebox.showerror("Error", f"Error reading sheets: {str(e)}")
    
    def validate_inputs(self):
        """Validate input parameters before processing."""
        adjusted_file = self.adjusted_rates_var.get()
        template_file = self.template_var.get()
        output_file = self.output_var.get()
        
        if not adjusted_file:
            messagebox.showerror("Error", "Please select an Adjusted Rates file.")
            return False
        
        if not template_file:
            messagebox.showerror("Error", "Please select a Template file.")
            return False
            
        if not output_file:
            messagebox.showerror("Error", "Please specify an output filename.")
            return False
        
        # Ensure output file has a valid Excel extension
        if not output_file.lower().endswith(('.xlsx', '.xls')):
            output_file = output_file + '.xlsx'
            self.output_var.set(output_file)
            logging.info(f"Added .xlsx extension to output file: {output_file}")
        
        # Validate input files exist
        if not os.path.exists(adjusted_file):
            messagebox.showerror("Error", f"Adjusted Rates file does not exist: {adjusted_file}")
            return False
            
        if not os.path.exists(template_file):
            messagebox.showerror("Error", f"Template file does not exist: {template_file}")
            return False
        
        # Validate input files are valid Excel files
        try:
            # Try to read at least one row from each file to validate format
            pd.read_excel(adjusted_file, nrows=1)
            pd.read_excel(template_file, nrows=1)
        except Exception as e:
            messagebox.showerror("Error", f"Invalid Excel file format: {str(e)}")
            return False
        
        # Check if output directory exists
        output_dir = os.path.dirname(output_file)
        if output_dir and not os.path.exists(output_dir):
            try:
                os.makedirs(output_dir)
                logging.info(f"Created output directory: {output_dir}")
            except Exception as e:
                messagebox.showerror("Error", f"Cannot create output directory: {str(e)}")
                return False
        
        # Check if we have write permissions in the output directory
        try:
            output_dir = output_dir if output_dir else '.'
            test_file = os.path.join(output_dir, '.test_write_permission')
            with open(test_file, 'w') as f:
                f.write('test')
            os.remove(test_file)
        except Exception as e:
            messagebox.showerror("Error", f"No write permission in output directory: {str(e)}")
            return False
        
        # Check if output file is open
        try:
            if os.path.exists(output_file):
                with open(output_file, 'a') as f:
                    pass
        except Exception as e:
            messagebox.showerror("Error", f"Cannot access output file. It may be open in another program: {str(e)}")
            return False
            
        return True
    
    def process_files(self):
        """Process the files with current settings."""
        if not self.validate_inputs():
            return
        
        # Save current settings
        self.save_settings()
        
        # Start processing in a separate thread
        self.status_var.set("Processing...")
        self.progress_var.set(0)
        
        # Disable buttons during processing
        self.disable_controls()
        
        # Start worker thread
        worker_thread = threading.Thread(target=self.process_files_worker)
        worker_thread.daemon = True
        worker_thread.start()
    
    def process_files_worker(self):
        """Worker thread for file processing."""
        try:
            adjusted_file = self.adjusted_rates_var.get()
            template_file = self.template_var.get()
            output_file = self.output_var.get()
            adjusted_sheet = self.adjusted_sheet_var.get()
            template_sheet = self.template_sheet_var.get()
            
            # Update progress - Step 1: Loading files
            self.update_status("Loading and analyzing template file...", 5)
            
            # First analyze the template file to determine required fields
            template_structure = self.file_analyzer.analyze_file_structure(
                template_file, template_sheet)
            
            # Extract template columns
            template_df = pd.read_excel(template_file, sheet_name=template_sheet)
            template_columns = template_df.columns.tolist()
            
            # Step 2: Building required fields
            self.update_status("Analyzing template structure...", 10)
            
            # Build dynamic required fields from template columns
            required_fields = self.extract_required_fields_from_template(template_columns)
            
            # Update mapping system with dynamic required fields
            self.mapping_system.set_required_fields(required_fields)
            
            # Step 3: Loading adjusted rates file
            self.update_status("Loading adjusted rates file...", 15)
            
            # Analyze adjusted rates file
            adjusted_structure = self.file_analyzer.analyze_file_structure(
                adjusted_file, adjusted_sheet)
            
            # Get column names
            if isinstance(adjusted_structure, dict) and 'columns' in adjusted_structure:
                source_columns = list(adjusted_structure['columns'].keys())
            else:
                # Read the file directly to get column names
                adjusted_df = pd.read_excel(adjusted_file, sheet_name=adjusted_sheet)
                source_columns = adjusted_df.columns.tolist()
            
            # Step 4: Generating mapping - more granular progress
            self.update_status(f"Generating column mapping for {len(source_columns)} columns...", 20)
            
            # Generate mapping
            use_saved = self.use_saved_var.get()
            use_enhanced = self.enhanced_format_var.get()
            
            # If using enhanced detection, modify the structure to use additional heuristics
            if use_enhanced:
                adjusted_structure['use_enhanced_detection'] = True
            
            mapping = self.mapping_system.generate_mapping(
                source_columns, use_saved_mappings=use_saved)
            
            # Step 5: Auto-detecting pivot columns
            self.update_status("Auto-detecting deductible and rate columns...", 25)
            
            # Add special handling for Deductible and RateCost columns
            # We need these for pivoting but they're not part of the required mapping fields
            self.detect_pivot_columns(source_columns, mapping, adjusted_structure)
            
            # Step 6: Check mapping confidence
            self.update_status("Validating mapping confidence...", 30)
            
            low_confidence = False
            mapped_count = len(mapping)
            required_count = len(required_fields)
            
            self.update_status(f"Mapping confidence: {mapped_count}/{required_count} fields mapped", 35)
            
            for field, conf in self.mapping_system.mapping_confidence.items():
                if conf < 70:  # Threshold for low confidence
                    low_confidence = True
                    break
                    
            # If low confidence and auto-detect is enabled, show mapping dialog
            if low_confidence and self.auto_detect_var.get():
                self.update_status("Low confidence mapping - awaiting user input...", 40)
                
                # Put request for mapping dialog in queue
                self.msg_queue.put(("show_mapping_dialog", {
                    "source_columns": source_columns,
                    "mapping": mapping,
                    "required_fields": required_fields
                }))
                
                # Wait for mapping update
                return
            
            # Continue with processing
            self.continue_processing(adjusted_file, template_file, output_file, 
                                    adjusted_sheet, template_sheet, mapping)
            
        except Exception as e:
            logging.error(f"Error in processing: {str(e)}", exc_info=True)
            self.update_status(f"Error: {str(e)}", 0)
            
            # Show error dialog
            self.msg_queue.put(("show_error", {
                "title": "Processing Error",
                "message": f"An error occurred during processing: {str(e)}"
            }))
        finally:
            # Re-enable controls
            self.msg_queue.put(("enable_controls", {}))
    
    def continue_processing(self, adjusted_file, template_file, output_file, 
                           adjusted_sheet, template_sheet, mapping):
        """Continue processing after mapping is confirmed."""
        try:
            # Step 7: Preparing for data processing
            self.update_status("Preparing for data transformation...", 45)
            
            # Make sure data processor has the default deductible value
            default_deductible = self.config_mgr.get_setting("default_deductible", "100")
            self.data_processor.default_deductible = default_deductible
            logging.info(f"Using default deductible for processing: {default_deductible}")
            
            # Ensure Deductible and RateCost columns are in the mapping
            if "Deductible" not in mapping or "RateCost" not in mapping:
                self.update_status("Detecting required pivot columns...", 50)
                # Try to detect them one more time from the source data
                adjusted_df = self.data_processor.load_excel_file(adjusted_file, adjusted_sheet)
                adjusted_structure = self.file_analyzer.analyze_file_structure(adjusted_file, adjusted_sheet)
                self.detect_pivot_columns(adjusted_df.columns.tolist(), mapping, adjusted_structure)
                
                # Log the results of pivot column detection
                if "Deductible" in mapping and "RateCost" in mapping:
                    logging.info(f"Successfully detected pivot columns: Deductible={mapping['Deductible']}, RateCost={mapping['RateCost']}")
                    self.update_status(f"Found pivot columns: Deductible={mapping['Deductible']}, RateCost={mapping['RateCost']}", 52)
                else:
                    logging.warning("Failed to detect Deductible and/or RateCost columns for pivoting.")
                    self.update_status("Warning: Could not detect all pivot columns", 52)
            
            # Step 8: Loading data with progress updates
            self.update_status("Loading source data...", 55)
            adjusted_df = self.data_processor.load_excel_file(adjusted_file, adjusted_sheet)
            
            self.update_status("Loading template data...", 58)
            template_df = self.data_processor.load_excel_file(template_file, template_sheet)
            
            # Check if we have data
            if adjusted_df.empty:
                self.update_status("Error: Adjusted rates file contains no data", 0)
                self.msg_queue.put(("show_error", {
                    "title": "Empty Data",
                    "message": "The adjusted rates file contains no data to process."
                }))
                return
            
            # Log information about mapping
            logging.info(f"Using mapping: {mapping}")
            logging.info(f"Adjusted dataframe columns: {adjusted_df.columns.tolist()}")
            logging.info(f"Adjusted dataframe shape: {adjusted_df.shape}")
            
            # Step 9: Transforming data with detailed progress
            row_count = len(adjusted_df)
            self.update_status(f"Transforming data ({row_count} rows)...", 60)
            
            # Before calling transform_data, add detailed logging for Deductible and RateCost
            if "Deductible" in mapping and "RateCost" in mapping:
                logging.info(f"Pivot columns found in mapping before transformation:")
                logging.info(f"  Deductible column: {mapping['Deductible']}")
                logging.info(f"  RateCost column: {mapping['RateCost']}")
                
                # Verify these columns exist in the dataframe
                if mapping['Deductible'] in adjusted_df.columns and mapping['RateCost'] in adjusted_df.columns:
                    logging.info("Both pivot columns found in the dataframe, proceeding with transformation")
                    self.update_status("Pivot columns verified, transforming data...", 65)
                else:
                    missing = []
                    if mapping['Deductible'] not in adjusted_df.columns:
                        missing.append(f"Deductible ({mapping['Deductible']})")
                    if mapping['RateCost'] not in adjusted_df.columns:
                        missing.append(f"RateCost ({mapping['RateCost']})")
                    logging.warning(f"Missing pivot columns in dataframe: {', '.join(missing)}")
                    self.update_status(f"Warning: Missing columns: {', '.join(missing)}", 65)
            else:
                logging.warning("Deductible and/or RateCost not found in mapping before transformation")
                self.update_status("Warning: Missing deductible/rate mapping", 65)
            
            # Step 10: Performing data transformation
            self.update_status("Applying column mapping and transforming data...", 70)
            
            # Transform data directly with adjusted dataframe and mapping
            transformed_df = self.data_processor.transform_data(adjusted_df, mapping)
            
            # Check if transformation returned data
            if transformed_df.empty:
                self.update_status("Error: No data after transformation", 0)
                self.msg_queue.put(("show_error", {
                    "title": "Transformation Error",
                    "message": "The data transformation process resulted in no data. Check the logs for details."
                }))
                return
            
            logging.info(f"Transformed data shape: {transformed_df.shape}")
            output_row_count = len(transformed_df)
            
            # Step 11: Integrating with template
            self.update_status(f"Integrating with template ({output_row_count} transformed rows)...", 80)
            
            # Integrate with template
            final_df = self.data_processor.integrate_with_template(transformed_df, template_df)
            
            # Check if final data is empty
            if final_df.empty and not transformed_df.empty:
                self.update_status("Warning: Template integration produced no data, using transformed data", 85)
                self.msg_queue.put(("show_error", {
                    "title": "Integration Warning",
                    "message": "The template integration process resulted in no data. Using transformed data instead."
                }))
                # Use transformed data if integration failed
                final_df = transformed_df
            
            final_row_count = len(final_df)
            logging.info(f"Final data shape: {final_df.shape} with {final_row_count} rows")
            
            # Step 12: Saving output with detailed progress
            self.update_status(f"Saving output file with {final_row_count} rows...", 90)
            
            # Make sure output file has the correct extension
            if not output_file.lower().endswith(('.xlsx', '.xls')):
                output_file = output_file + '.xlsx'
                logging.info(f"Added .xlsx extension to output file: {output_file}")
            
            try:
                # Use the new save_excel_file method from DataProcessor
                self.data_processor.save_excel_file(final_df, output_file, sheet_name=template_sheet)
                success = True
            except Exception as e:
                error_msg = f"Error saving file: {str(e)}"
                logging.error(error_msg)
                self.update_status(error_msg, 0)
                self.msg_queue.put(("show_error", {
                    "title": "Save Error",
                    "message": error_msg
                }))
                return
            
            # Finish processing with success message
            self.update_status(f"Processing complete! Created file with {final_row_count} rows", 100)
            
            # Show success message
            self.msg_queue.put(("show_info", {
                "title": "Processing Complete",
                "message": f"Successfully processed {row_count} rows and created output file with {final_row_count} rows."
            }))
            
            # Open the file if requested
            if self.auto_open_var.get():
                self.msg_queue.put(("open_file", {
                    "file_path": output_file
                }))
                
        except Exception as e:
            logging.error(f"Error in continue_processing: {str(e)}", exc_info=True)
            self.update_status(f"Error: {str(e)}", 0)
            self.msg_queue.put(("show_error", {
                "title": "Processing Error",
                "message": f"Error during processing: {str(e)}"
            }))
        finally:
            # Re-enable controls
            self.msg_queue.put(("enable_controls", {}))
    
    def preview_mapping(self):
        """Preview the column mapping with dynamically detected template columns."""
        if not self.validate_inputs():
            return
        
        adjusted_file = self.adjusted_rates_var.get()
        adjusted_sheet = self.adjusted_sheet_var.get()
        template_file = self.template_var.get()
        template_sheet = self.template_sheet_var.get()
        
        try:
            # First load the template file to extract columns
            logging.info(f"Loading template file: {template_file}, sheet: {template_sheet}")
            template_df = pd.read_excel(template_file, sheet_name=template_sheet)
            
            # Get all columns from template
            template_columns = template_df.columns.tolist()
            logging.info(f"Found {len(template_columns)} columns in template: {template_columns}")
            
            # Build dynamic required fields from template columns
            required_fields = self.extract_required_fields_from_template(template_columns)
            logging.info(f"Extracted required fields: {required_fields}")
            
            # Update mapping system with dynamic required fields
            self.mapping_system.set_required_fields(required_fields)
            
            # Load the adjusted rates file
            logging.info(f"Loading adjusted rates file: {adjusted_file}, sheet: {adjusted_sheet}")
            adjusted_df = pd.read_excel(adjusted_file, sheet_name=adjusted_sheet)
            source_columns = adjusted_df.columns.tolist()
            logging.info(f"Found {len(source_columns)} columns in adjusted rates file")
            
            # Analyze the source file for structure
            adjusted_structure = self.file_analyzer.analyze_file_structure(
                adjusted_file, adjusted_sheet)
            
            # Generate mapping suggestions 
            use_saved = self.use_saved_var.get()
            mapping = self.mapping_system.generate_mapping(source_columns, use_saved_mappings=use_saved)
            logging.info(f"Generated mapping with {len(mapping)} fields mapped")
            
            # Add special handling for Deductible and RateCost columns
            # These should be identified automatically but not shown in the mapping dialog
            self.detect_pivot_columns(source_columns, mapping, adjusted_structure)
            
            # Add a note to explain the special handling
            if "Deductible" in mapping and "RateCost" in mapping:
                deduct_col = mapping["Deductible"]
                rate_col = mapping["RateCost"]
                special_note = (f"Note: The columns '{deduct_col}' and '{rate_col}' will be used to populate "
                               f"the deductible columns (Deduct0, Deduct50, etc.) in the template. "
                               f"They are handled automatically and don't need to be mapped.")
                logging.info(special_note)
                
                # Show this explanation to the user before showing mapping dialog
                messagebox.showinfo("Special Column Handling", special_note)
            
            # Show mapping dialog with dynamic fields
            self.show_mapping_dialog(
                source_columns=source_columns,
                mapping=mapping,
                required_fields=required_fields
            )
            
        except Exception as e:
            logging.error(f"Error previewing mapping: {str(e)}", exc_info=True)
            messagebox.showerror("Error", f"An error occurred while previewing mapping: {str(e)}")

    def extract_required_fields_from_template(self, template_columns):
        """
        Extract required fields from template columns.
        
        Args:
            template_columns: List or pandas Index of column names from template
            
        Returns:
            list: Required fields for mapping
        """
        required_fields = []
        deductible_columns = []
        
        # Ensure template_columns is a list
        if not isinstance(template_columns, list):
            template_columns = list(template_columns)
        
        logging.info(f"Extracting required fields from {len(template_columns)} template columns")
        
        # Define essential columns that should always be included if present
        essential_columns = ['CompanyCode', 'Term', 'Miles', 'FromMiles', 'ToMiles', 
                           'Coverage', 'State', 'Class', 'PlanDeduct']
        
        # Process each column from template
        for col in template_columns:
            col_str = str(col).lower()
            
            # Special handling for deductible columns (Deduct0, Deduct50, etc.)
            if col_str.startswith('deduct'):
                deductible_columns.append(col)
                continue
            
            # Skip PlanDeduct as it's handled separately
            if col_str == 'plandeduct':
                continue
            
            # Add all other template columns as required fields
            required_fields.append(col)
        
        # Ensure essential columns are included in the required fields
        for essential_col in essential_columns:
            if essential_col not in required_fields and essential_col.lower() not in [f.lower() for f in required_fields]:
                required_fields.append(essential_col)
        
        # Log the fields we extracted
        logging.info(f"Extracted {len(required_fields)} required fields from template")
        logging.info(f"Found {len(deductible_columns)} deductible columns that will be auto-populated")
        
        return required_fields

    def show_mapping_dialog(self, source_columns, mapping, required_fields=None):
        """
        Show the mapping dialog to confirm column mappings.
        
        Args:
            source_columns (list): List of column names from source file
            mapping (dict): Dictionary of field -> column mappings
            required_fields (list, optional): List of required fields, defaults to mapping system's required fields
        """
        try:
            # Log initial state of parameters
            logging.info(f"Show mapping dialog called with: {len(source_columns)} source columns, " 
                        f"{len(mapping) if mapping else 0} mappings, "
                        f"required_fields provided: {required_fields is not None}")
            
            # Validate source_columns is a list
            if not isinstance(source_columns, list):
                logging.warning(f"source_columns is not a list, converting from {type(source_columns)}")
                if isinstance(source_columns, (pd.Index, pd.Series)):
                    source_columns = source_columns.tolist()
                else:
                    source_columns = list(source_columns) if hasattr(source_columns, '__iter__') else []
            
            # Validate mapping is a dictionary
            if not isinstance(mapping, dict):
                logging.warning(f"mapping is not a dictionary, creating empty dict from {type(mapping)}")
                mapping = {}
            
            # Use provided required fields or fall back to default
            if required_fields is None:
                logging.info("No required_fields provided, using mapping system's required fields")
                required_fields = self.mapping_system.required_fields
            
            # Ensure required_fields is a list
            if not isinstance(required_fields, list):
                logging.warning(f"required_fields is not a list, converting from {type(required_fields)}")
                if isinstance(required_fields, (pd.Index, pd.Series)):
                    required_fields = required_fields.tolist()
                else:
                    required_fields = list(required_fields) if hasattr(required_fields, '__iter__') else []
            
            logging.info(f"Showing mapping dialog with {len(required_fields)} required fields")
            logging.info(f"Source columns (first 5): {source_columns[:5] if len(source_columns) > 5 else source_columns}")
            
            # Create dialog with validated parameters
            dialog = MappingDialog(
                self, 
                source_columns=source_columns,
                required_fields=required_fields,
                suggested_mapping=mapping
            )
            
            # Show dialog and get result
            result_mapping, save_as_template, template_name = dialog.show()
            
            if result_mapping:
                # Update mapping
                self.mapping_system.current_mapping = result_mapping
                
                # Get default deductible from the UI field
                default_deductible = self.default_deduct_var.get()
                self.config_mgr.set_setting("default_deductible", default_deductible)
                self.data_processor.default_deductible = default_deductible
                logging.info(f"Using default deductible from UI: {default_deductible}")
                
                # Save mapping if requested
                if save_as_template and template_name:
                    self.mapping_system.save_current_mapping(source_columns, template_name)
                    messagebox.showinfo("Mapping Saved", f"Mapping template '{template_name}' has been saved.")
                
                # If called from worker thread, continue processing
                if self.status_var.get() == "Low confidence mapping - awaiting user input...":
                    # Continue processing in a new thread
                    worker_thread = threading.Thread(
                        target=self.continue_processing,
                        args=(
                            self.adjusted_rates_var.get(),
                            self.template_var.get(),
                            self.output_var.get(),
                            self.adjusted_sheet_var.get(),
                            self.template_sheet_var.get(),
                            result_mapping
                        )
                    )
                    worker_thread.daemon = True
                    worker_thread.start()
        except Exception as e:
            logging.error(f"Error showing mapping dialog: {str(e)}", exc_info=True)
            messagebox.showerror("Error", f"Failed to show mapping dialog: {str(e)}")

    def show_mappings_manager(self):
        """Show the dialog to manage saved mappings."""
        # TO DO: Implement mappings manager dialog
        messagebox.showinfo("Not Implemented", "Mapping Manager will be implemented in a future version.")
    
    def show_advanced_options(self):
        """Show advanced options dialog."""
        # TO DO: Implement advanced options dialog
        messagebox.showinfo("Not Implemented", "Advanced Options will be implemented in a future version.")
    
    def update_status(self, message, progress):
        """
        Update status message and progress bar.
        
        Args:
            message: Status message to display
            progress: Progress value (0-100)
        """
        # Update status via queue for thread safety
        self.msg_queue.put(("update_status", {
            "message": message,
            "progress": progress
        }))
        
        # Also update progress percentage
        percentage = f"{int(progress)}%"
        self.msg_queue.put(("update_progress_percentage", {
            "percentage": percentage
        }))
    
    def process_queue(self):
        """Process messages from the queue."""
        try:
            while True:
                action, params = self.msg_queue.get_nowait()
                
                if action == "update_status":
                    self.status_var.set(params["message"])
                    self.progress_var.set(params["progress"])
                    
                elif action == "update_progress_percentage":
                    self.progress_percentage.set(params["percentage"])
                    
                elif action == "show_error":
                    messagebox.showerror(params["title"], params["message"])
                    
                elif action == "show_info":
                    messagebox.showinfo(params["title"], params["message"])
                    
                elif action == "enable_controls":
                    self.enable_controls()
                    
                elif action == "open_file":
                    self.open_file(params["file_path"])
                    
                elif action == "show_mapping_dialog":
                    try:
                        # Initialize with defaults
                        source_columns = []
                        mapping = {}
                        required_fields = None
                        
                        # Handle different param formats
                        if isinstance(params, list):
                            # Unpack list params based on position
                            # Expected format: [source_columns, mapping, required_fields]
                            if len(params) > 0:
                                source_columns = params[0]
                            if len(params) > 1:
                                mapping = params[1]
                            if len(params) > 2:
                                required_fields = params[2]
                            logging.info(f"Received list params for mapping dialog: {len(params)} items")
                        elif isinstance(params, dict):
                            # Extract from dictionary
                            source_columns = params.get("source_columns", [])
                            mapping = params.get("mapping", {})
                            required_fields = params.get("required_fields", None)
                            logging.info(f"Received dict params for mapping dialog with keys: {list(params.keys())}")
                        else:
                            logging.warning(f"Unexpected params type: {type(params)}")
                        
                        # Ensure source_columns is a list
                        if not isinstance(source_columns, list):
                            logging.info(f"Converting source_columns from {type(source_columns)} to list")
                            if isinstance(source_columns, (pd.Index, pd.Series)):
                                source_columns = source_columns.tolist()
                            else:
                                source_columns = list(source_columns) if hasattr(source_columns, '__iter__') else []
                        
                        # Ensure mapping is a dictionary
                        if not isinstance(mapping, dict):
                            logging.warning(f"mapping is not a dictionary: {type(mapping)}, creating empty dict")
                            mapping = {}
                        
                        logging.info(f"Showing mapping dialog with {len(source_columns)} source columns and {len(mapping)} mappings")
                        
                        # Call the method with properly formatted parameters
                        self.show_mapping_dialog(
                            source_columns=source_columns,
                            mapping=mapping,
                            required_fields=required_fields
                        )
                    except Exception as e:
                        logging.error(f"Error in queue handler for show_mapping_dialog: {str(e)}", exc_info=True)
                        messagebox.showerror("Error", f"Failed to show mapping dialog: {str(e)}")
                
                self.msg_queue.task_done()
                
        except queue.Empty:
            pass
            
        self.after(100, self.process_queue)
    
    def disable_controls(self):
        """Disable controls during processing."""
        for child in self.winfo_children():
            for widget in child.winfo_children():
                if isinstance(widget, ttk.Button):
                    widget.state(['disabled'])
    
    def enable_controls(self):
        """Enable controls after processing."""
        for child in self.winfo_children():
            for widget in child.winfo_children():
                if isinstance(widget, ttk.Button):
                    widget.state(['!disabled'])
    
    def open_file(self, file_path):
        """Open a file with the default application."""
        try:
            # Check if file exists first to avoid error
            if not os.path.exists(file_path):
                logging.warning(f"File does not exist, cannot open: {file_path}")
                messagebox.showwarning("File Not Found", 
                                     f"The file could not be opened because it doesn't exist:\n\n{file_path}")
                return
                
            # Use the appropriate method based on the platform
            if sys.platform.startswith('win'):
                os.startfile(file_path)
            elif sys.platform.startswith('darwin'):  # macOS
                subprocess.call(['open', file_path])
            else:  # Linux
                subprocess.call(['xdg-open', file_path])
                
            logging.info(f"Opened file: {file_path}")
            
        except Exception as e:
            logging.error(f"Error opening file: {str(e)}")
            messagebox.showerror("Error", f"Could not open file: {str(e)}")
    
    def on_exit(self):
        """Handle application exit."""
        # Save settings
        self.save_settings()
        
        logging.info("Application exiting")
        self.destroy()

    def analyze_files(self):
        """Analyze the selected files for format detection without processing."""
        adjusted_file = self.adjusted_rates_var.get()
        adjusted_sheet = self.adjusted_sheet_var.get()
        template_file = self.template_var.get()
        template_sheet = self.template_sheet_var.get()
        
        if not adjusted_file or not os.path.exists(adjusted_file):
            messagebox.showerror("Error", "Please select a valid Adjusted Rates file.")
            return
        
        try:
            # Update status
            self.status_var.set("Analyzing files...")
            self.progress_var.set(10)
            
            # Analyze adjusted rates file
            adjusted_structure = self.file_analyzer.analyze_file_structure(
                adjusted_file, adjusted_sheet)
            
            # Extract key information from the analysis
            col_count = adjusted_structure.get('column_count', 0)
            row_count = adjusted_structure.get('row_count', 0)
            patterns = adjusted_structure.get('patterns', {})
            
            # Update format detection status
            format_msg = f"Adjusted Rates: {col_count} columns, {row_count} rows."
            
            # Add info about deductible pattern if detected
            if patterns.get('has_deductible_data', False):
                deduct_pattern = patterns.get('pattern', 'unknown')
                deduct_values = patterns.get('values', [])
                format_msg += f" Deductibles: {deduct_pattern} ({', '.join(map(str, deduct_values))})"
            
            self.format_detection_var.set(format_msg)
            
            # If template file is provided, analyze it too
            if template_file and os.path.exists(template_file):
                self.progress_var.set(30)
                template_structure = self.file_analyzer.analyze_file_structure(
                    template_file, template_sheet)
                
                template_col_count = template_structure.get('column_count', 0)
                template_row_count = template_structure.get('row_count', 0)
                
                # Update status with template info
                self.status_var.set(f"Analyzed both files. Template: {template_col_count} columns, {template_row_count} rows.")
            
            # Generate mapping
            self.progress_var.set(50)
            use_saved = self.use_saved_var.get()
            use_enhanced = self.enhanced_format_var.get()
            
            # If using enhanced detection, modify the structure to use additional heuristics
            if use_enhanced:
                adjusted_structure['use_enhanced_detection'] = True
            
            # Get column names as a list
            if isinstance(adjusted_structure, dict) and 'columns' in adjusted_structure:
                source_columns = list(adjusted_structure['columns'].keys())
            else:
                # Read the file directly to get column names
                adjusted_df = pd.read_excel(adjusted_file, sheet_name=adjusted_sheet)
                source_columns = adjusted_df.columns.tolist()
            
            mapping = self.mapping_system.generate_mapping(
                source_columns, use_saved_mappings=use_saved)
            
            # Update mapping confidence indicator
            if self.mapping_system.mapping_confidence:
                # Calculate average confidence
                confidence_values = list(self.mapping_system.mapping_confidence.values())
                if confidence_values:
                    avg_confidence = sum(confidence_values) / len(confidence_values)
                    self.mapping_confidence_var.set(int(avg_confidence))
                    self.mapping_confidence_label.config(text=f"{int(avg_confidence)}%")
                    
                    # Color code based on confidence
                    if avg_confidence < 60:
                        self.mapping_confidence_label.config(foreground="red")
                    elif avg_confidence < 80:
                        self.mapping_confidence_label.config(foreground="orange")
                    else:
                        self.mapping_confidence_label.config(foreground="green")
            
            # Show a summary
            self.progress_var.set(100)
            
            # Count mapped fields
            mapped_count = len(mapping)
            required_count = len(self.mapping_system.required_fields)
            
            # Show mapping analysis dialog
            message = f"File Analysis Complete\n\n"
            message += f"Adjusted Rates: {col_count} columns, {row_count} rows\n"
            message += f"Mapping: {mapped_count} of {required_count} required fields mapped\n\n"
            
            # Add warning for low confidence
            low_confidence = any(conf < 70 for conf in self.mapping_system.mapping_confidence.values())
            if low_confidence:
                message += "⚠️ Some fields have low mapping confidence.\n"
                message += "You may need to manually map columns during processing.\n\n"
            
            # Add message about proceeding
            message += "Do you want to see the current mapping details?"
            
            if messagebox.askyesno("Analysis Complete", message):
                # Get column names as a list for the dialog
                if isinstance(adjusted_structure, dict) and 'columns' in adjusted_structure:
                    # Extract column names from structure
                    source_columns = list(adjusted_structure['columns'].keys())
                    logging.info(f"Extracted {len(source_columns)} columns from structure for mapping dialog")
                else:
                    # Read the file directly to get column names if needed
                    logging.info(f"Reading source columns directly from file for mapping dialog")
                    adjusted_df = pd.read_excel(adjusted_file, sheet_name=adjusted_sheet)
                    source_columns = adjusted_df.columns.tolist()
                
                # Show mapping dialog with correct parameters
                self.show_mapping_dialog(
                    source_columns=source_columns,
                    mapping=mapping,
                    required_fields=self.mapping_system.required_fields
                )
            
            self.status_var.set("File analysis complete")
            
        except Exception as e:
            logging.error(f"Error analyzing files: {str(e)}", exc_info=True)
            self.status_var.set(f"Error: {str(e)}")
            self.progress_var.set(0)
            messagebox.showerror("Error", f"Error analyzing files: {str(e)}")

    def setup_styles(self):
        """Set up the ttk styles for the modern Moxy theme."""
        style = ttk.Style()
        
        # Set theme to 'clam' for better styling control
        if style.theme_use() != 'clam':
            style.theme_use('clam')
        
        # Configure base styles with modern look
        style.configure("TFrame", background=self.DARK_BG)
        
        # LabelFrame with subtle border and darker background
        style.configure("TLabelframe", 
                       background=self.DARKER_BG,
                       bordercolor=self.BORDER_COLOR,
                       darkcolor=self.BORDER_COLOR,
                       lightcolor=self.BORDER_COLOR)
        style.configure("TLabelframe.Label", 
                       background=self.DARK_BG,
                       foreground="#FFFFFF",
                       font=("Segoe UI", 10, "bold"))
        
        # Modern button style with rounded corners and gradient effect
        style.configure("TButton",
                       background=self.BUTTON_BG,
                       foreground="#FFFFFF",
                       font=("Segoe UI", 9, "bold"),
                       padding=(20, 10),
                       borderwidth=0,
                       relief="flat")
        style.map("TButton",
                 background=[("active", self.BUTTON_HOVER_BG),
                           ("pressed", self.BUTTON_ACTIVE_BG),
                           ("disabled", self.BUTTON_DISABLED_BG)],
                 foreground=[("disabled", "#999999")],
                 relief=[("pressed", "sunken"),
                        ("active", "flat")])
        
        # Dark button style
        style.configure("Dark.TButton",
                       background=self.BUTTON_BG,
                       foreground="#FFFFFF",
                       font=("Segoe UI", 9, "bold"),
                       padding=(20, 10),
                       borderwidth=1,
                       relief="solid")
        style.map("Dark.TButton",
                 background=[("active", self.BUTTON_HOVER_BG),
                           ("pressed", self.BUTTON_ACTIVE_BG),
                           ("disabled", self.BUTTON_DISABLED_BG)],
                 foreground=[("disabled", "#999999")],
                 relief=[("pressed", "sunken"),
                        ("active", "solid")])
        
        # Process button with accent color
        style.configure("Process.TButton",
                       background=self.ACCENT_COLOR,
                       foreground="#000000",
                       font=("Segoe UI", 9, "bold"),
                       padding=(20, 10),
                       borderwidth=1,
                       relief="solid")
        style.map("Process.TButton",
                 background=[("active", "#9ED84F"),
                           ("pressed", "#7AB52F"),
                           ("disabled", self.BUTTON_DISABLED_BG)],
                 foreground=[("disabled", "#999999")],
                 relief=[("pressed", "sunken"),
                        ("active", "solid")])
        
        # Entry fields with darker background for better contrast
        style.configure("TEntry",
                       background=self.DARKER_BG,
                       fieldbackground=self.DARKER_BG,
                       foreground="#FFFFFF",
                       insertcolor="#FFFFFF",
                       borderwidth=1,
                       relief="solid",
                       padding=8)
        style.map("TEntry",
                 fieldbackground=[("disabled", self.DARK_BG),
                                ("readonly", self.DARKER_BG)],
                 background=[("disabled", self.DARK_BG),
                           ("readonly", self.DARKER_BG)],
                 bordercolor=[("focus", self.ACCENT_COLOR)])
        
        # Labels with clean font and bright text
        style.configure("TLabel",
                       background=self.DARK_BG,
                       foreground="#FFFFFF",
                       font=("Segoe UI", 9))
        
        # Checkbuttons with modern look
        style.configure("TCheckbutton",
                       background=self.DARK_BG,
                       foreground="#FFFFFF",
                       font=("Segoe UI", 9))
        style.map("TCheckbutton",
                 background=[("active", self.DARK_BG)],
                 foreground=[("disabled", "#999999")])
        
        # Dark Combobox style
        style.configure("Dark.TCombobox",
                       fieldbackground=self.DARKER_BG,
                       background=self.DARKER_BG,
                       foreground="#FFFFFF",
                       arrowcolor="#FFFFFF",
                       selectbackground=self.BUTTON_BG,
                       selectforeground="#FFFFFF",
                       borderwidth=1,
                       padding=8)
        style.map("Dark.TCombobox",
                 fieldbackground=[("readonly", self.DARKER_BG),
                                ("disabled", self.DARK_BG),
                                ("active", self.DARKER_BG),
                                ("focus", self.DARKER_BG)],
                 background=[("readonly", self.DARKER_BG),
                           ("disabled", self.DARK_BG),
                           ("active", self.DARKER_BG),
                           ("focus", self.DARKER_BG)],
                 foreground=[("disabled", "#999999")],
                 selectbackground=[("readonly", self.BUTTON_BG)])
        
        # Configure base TCombobox style to match Dark.TCombobox
        style.configure("TCombobox",
                       fieldbackground=self.DARKER_BG,
                       background=self.DARKER_BG,
                       foreground="#FFFFFF",
                       arrowcolor="#FFFFFF",
                       selectbackground=self.BUTTON_BG,
                       selectforeground="#FFFFFF",
                       borderwidth=1,
                       padding=8)
        style.map("TCombobox",
                 fieldbackground=[("readonly", self.DARKER_BG),
                                ("disabled", self.DARK_BG),
                                ("active", self.DARKER_BG),
                                ("focus", self.DARKER_BG)],
                 background=[("readonly", self.DARKER_BG),
                           ("disabled", self.DARK_BG),
                           ("active", self.DARKER_BG),
                           ("focus", self.DARKER_BG)],
                 foreground=[("disabled", "#999999")],
                 selectbackground=[("readonly", self.BUTTON_BG)])
        
        # Progress bar with Moxy colors and increased height
        style.configure("Horizontal.TProgressbar",
                       background=self.PROGRESS_BG,
                       troughcolor=self.PROGRESS_TROUGH,
                       bordercolor=self.PROGRESS_TROUGH,
                       lightcolor=self.PROGRESS_BG,
                       darkcolor=self.PROGRESS_BG,
                       thickness=12)
        
        # Scrollbar with modern look
        style.configure("TScrollbar",
                       background=self.BUTTON_BG,
                       troughcolor=self.DARKER_BG,
                       borderwidth=0,
                       arrowcolor="#FFFFFF",
                       relief="flat")
        style.map("TScrollbar",
                 background=[("active", self.BUTTON_HOVER_BG),
                           ("pressed", self.BUTTON_ACTIVE_BG)])
        
        # Status section styles with darker background
        style.configure("Status.TFrame",
                       background=self.DARKER_BG,
                       relief="flat")
        style.configure("Status.TLabel",
                       background=self.DARKER_BG,
                       foreground="#FFFFFF",
                       font=("Segoe UI", 9))
        style.configure("Subheader.TLabel",
                       background=self.DARKER_BG,
                       foreground="#FFFFFF",
                       font=("Segoe UI", 9, "bold"))
        
        # Error label style
        style.configure("Error.TLabel",
                       background=self.DARK_BG,
                       foreground="#FF4444",
                       font=("Segoe UI", 9))
        
        # Configure the main window
        self.configure(background=self.DARK_BG)
        
        # Add padding and spacing to all frames
        style.configure("Padded.TFrame",
                       background=self.DARK_BG,
                       padding=10)
        
        # Modern tooltip style
        style.configure("Tooltip.TLabel",
                       background=self.DARKER_BG,
                       foreground="#FFFFFF",
                       font=("Segoe UI", 9),
                       padding=8)
        
        # Separator style
        style.configure("Separator.TFrame",
                       background=self.BORDER_COLOR)

    def detect_pivot_columns(self, source_columns, mapping, adjusted_structure):
        """
        Detect and add Deductible and RateCost columns to the mapping.
        
        Args:
            source_columns: List of column names from source file
            mapping: Dictionary of field -> column mappings
            adjusted_structure: Structure of the adjusted rates file
        """
        logging.info("Detecting pivot columns for Deductible and RateCost")
        
        # First check if they're already in the mapping
        has_deductible = "Deductible" in mapping
        has_rate_cost = "RateCost" in mapping
        
        if has_deductible and has_rate_cost:
            logging.info("Both Deductible and RateCost are already mapped")
            return
        
        # Try to find columns if they're not mapped yet
        # Search patterns for deductible columns
        deductible_patterns = ["deductible", "deduct", "ded", "deduc"]
        rate_cost_patterns = ["ratecost", "rate cost", "cost", "price", "premium", "rate"]
        
        if not has_deductible:
            # Look for deductible column
            for col in source_columns:
                col_lower = str(col).lower()
                if any(pattern in col_lower for pattern in deductible_patterns):
                    mapping["Deductible"] = col
                    logging.info(f"Auto-detected Deductible column: {col}")
                    has_deductible = True
                    break
                    
        if not has_rate_cost:
            # Look for rate cost column
            for col in source_columns:
                col_lower = str(col).lower()
                if any(pattern in col_lower for pattern in rate_cost_patterns):
                    mapping["RateCost"] = col
                    logging.info(f"Auto-detected RateCost column: {col}")
                    has_rate_cost = True
                    break
        
        # If we still haven't found them, try using structure analysis
        if (not has_deductible or not has_rate_cost) and isinstance(adjusted_structure, dict):
            columns_info = adjusted_structure.get('columns', {})
            patterns = adjusted_structure.get('patterns', {})
            
            # Check if file analysis found likely deductible column
            if not has_deductible and patterns.get('has_deductible_data', False):
                deduct_col = patterns.get('deductible_column')
                if deduct_col and deduct_col in source_columns:
                    mapping["Deductible"] = deduct_col
                    logging.info(f"Found Deductible column from structure analysis: {deduct_col}")
                    has_deductible = True
            
            # Check for likely rate cost column based on numeric analysis
            if not has_rate_cost:
                # Look for column with numeric values that might be costs
                number_cols = []
                for col, info in columns_info.items():
                    if info.get('data_type') == 'numeric' and col in source_columns:
                        number_cols.append(col)
                
                # If we have just one numeric column left, use it
                if len(number_cols) == 1:
                    mapping["RateCost"] = number_cols[0]
                    logging.info(f"Using single numeric column as RateCost: {number_cols[0]}")
                    has_rate_cost = True
                
                # Try to find based on column statistics
                elif len(number_cols) > 1:
                    # Look for columns with values that look like costs (decimals, reasonable ranges)
                    for col in number_cols:
                        col_info = columns_info.get(col, {})
                        min_val = col_info.get('min_value', 0)
                        max_val = col_info.get('max_value', 0)
                        
                        # Typical rate costs are positive and in a reasonable range
                        if min_val >= 0 and max_val < 10000:
                            mapping["RateCost"] = col
                            logging.info(f"Selected likely RateCost column based on value range: {col}")
                            has_rate_cost = True
                            break
        
        # Log error if we still couldn't find them
        if not has_deductible:
            logging.warning("Could not auto-detect Deductible column. User will need to specify it.")
        
        if not has_rate_cost:
            logging.warning("Could not auto-detect RateCost column. User will need to specify it.")

    def add_mapping_field(self, field):
        """
        Add a special field to the mapping system without requiring user mapping.
        
        Args:
            field: Field name to add (e.g., 'Deductible', 'RateCost')
        """
        logging.info(f"Adding special field to mapping system: {field}")
        # Nothing to do here since we're now handling these fields separately
        # in the detect_pivot_columns method
        pass

    def get_current_mapping_fields(self):
        """
        Get the current fields being mapped.
        
        Returns:
            list: List of field names 
        """
        if hasattr(self.mapping_system, 'current_mapping'):
            return list(self.mapping_system.current_mapping.keys())
        return []

    def _create_tooltip(self, widget, text):
        """Create a tooltip for a given widget."""
        tooltip = tk.Toplevel(widget)
        tooltip.withdraw()  # Hide initially
        tooltip.overrideredirect(True)  # Remove window decorations
        
        # Create tooltip content
        label = ttk.Label(tooltip, text=text, style="Dark.TLabel",
                         background=self.DARK_BG, foreground=self.TEXT_COLOR,
                         wraplength=300, padding=5)
        label.pack()
        
        def show_tooltip(event=None):
            """Show the tooltip."""
            tooltip.deiconify()
            # Position tooltip near the widget
            x = widget.winfo_rootx() + widget.winfo_width()
            y = widget.winfo_rooty()
            tooltip.geometry(f"+{x}+{y}")
        
        def hide_tooltip(event=None):
            """Hide the tooltip."""
            tooltip.withdraw()
        
        # Bind tooltip to widget events
        widget.bind("<Enter>", show_tooltip)
        widget.bind("<Leave>", hide_tooltip)
        widget.bind("<Button-1>", hide_tooltip)  # Hide on click
        
        return tooltip

def main():
    app = Application()
    app.mainloop()

if __name__ == "__main__":
    main() 