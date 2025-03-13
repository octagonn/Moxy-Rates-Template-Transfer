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
from tkinter import ttk, filedialog, messagebox
import configparser
import threading
import queue
import logging
from datetime import datetime

# Import custom modules
from file_analyzer import FileAnalyzer
from mapping_system import MappingSystem, MappingDialog
from config_manager import ConfigManager, MappingConfigManager
from data_processor import DataProcessor

class Application(tk.Tk):
    """Main application window for Moxy Rates Template Transfer."""
    
    def __init__(self):
        super().__init__()
        
        # Configure window
        self.title("Moxy Rates Template Transfer")
        self.geometry("800x600")
        self.minsize(800, 600)
        
        # Initialize configuration
        self.config_mgr = ConfigManager()
        self.mapping_config = MappingConfigManager()
        
        # Set up logging
        self.setup_logging()
        
        # Initialize components
        self.file_analyzer = FileAnalyzer()
        self.mapping_system = MappingSystem(self.mapping_config)
        self.data_processor = DataProcessor()
        
        # Create UI components
        self.create_ui()
        
        # Load saved settings
        self.load_settings()
        
        # Message queue for thread communication
        self.msg_queue = queue.Queue()
        self.after(100, self.process_queue)
        
        logging.info("Application initialized")
    
    def setup_logging(self):
        """Set up logging configuration."""
        log_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logs")
        if not os.path.exists(log_dir):
            os.makedirs(log_dir)
            
        log_file = os.path.join(log_dir, f"app_{datetime.now().strftime('%Y%m%d')}.log")
        
        # Configure logging
        log_level = logging.DEBUG if self.config_mgr.get_setting("enable_logging", False) else logging.INFO
        
        logging.basicConfig(
            level=log_level,
            format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(log_file),
                logging.StreamHandler()
            ]
        )
    
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
        """Create the input files section."""
        input_frame = ttk.LabelFrame(parent, text="Input Files", padding="10")
        input_frame.pack(fill=tk.X, pady=5)
        
        # Adjusted Rates file
        ttk.Label(input_frame, text="Adjusted Rates File:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.adjusted_rates_var = tk.StringVar()
        ttk.Entry(input_frame, textvariable=self.adjusted_rates_var, width=50).grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(input_frame, text="Browse...", command=self.browse_adjusted_rates).grid(row=0, column=2, padx=5, pady=5)
        
        # Template file
        ttk.Label(input_frame, text="Template File:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.template_var = tk.StringVar()
        ttk.Entry(input_frame, textvariable=self.template_var, width=50).grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(input_frame, text="Browse...", command=self.browse_template).grid(row=1, column=2, padx=5, pady=5)
    
    def create_output_section(self, parent):
        """Create the output section."""
        output_frame = ttk.LabelFrame(parent, text="Output", padding="10")
        output_frame.pack(fill=tk.X, pady=5)
        
        # Output file
        ttk.Label(output_frame, text="Output Filename:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.output_var = tk.StringVar()
        ttk.Entry(output_frame, textvariable=self.output_var, width=50).grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(output_frame, text="Browse...", command=self.browse_output).grid(row=0, column=2, padx=5, pady=5)
    
    def create_options_section(self, parent):
        """Create the options section."""
        options_frame = ttk.LabelFrame(parent, text="Options", padding="10")
        options_frame.pack(fill=tk.X, pady=5)
        
        # Checkboxes for options
        self.remember_dirs_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(options_frame, text="Remember last used directories", 
                     variable=self.remember_dirs_var).grid(row=0, column=0, sticky=tk.W, pady=2)
        
        self.open_after_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(options_frame, text="Open output file after processing", 
                     variable=self.open_after_var).grid(row=1, column=0, sticky=tk.W, pady=2)
        
        self.enable_logging_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(options_frame, text="Enable detailed logging", 
                     variable=self.enable_logging_var).grid(row=2, column=0, sticky=tk.W, pady=2)
        
        self.auto_detect_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(options_frame, text="Auto-detect file formats", 
                     variable=self.auto_detect_var).grid(row=0, column=1, sticky=tk.W, pady=2)
        
        self.use_saved_mappings_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(options_frame, text="Use saved mappings when available", 
                     variable=self.use_saved_mappings_var).grid(row=1, column=1, sticky=tk.W, pady=2)
        
        # Sheet selection
        ttk.Label(options_frame, text="Adjusted Rates Sheet:").grid(row=3, column=0, sticky=tk.W, pady=5)
        self.adjusted_sheet_var = tk.StringVar(value="Dealer Cost Rates")
        self.adjusted_sheet_combo = ttk.Combobox(options_frame, textvariable=self.adjusted_sheet_var, width=30)
        self.adjusted_sheet_combo.grid(row=3, column=1, sticky=tk.W, padx=5, pady=5)
        
        ttk.Label(options_frame, text="Template Sheet:").grid(row=4, column=0, sticky=tk.W, pady=5)
        self.template_sheet_var = tk.StringVar(value="Sheet1")
        self.template_sheet_combo = ttk.Combobox(options_frame, textvariable=self.template_sheet_var, width=30)
        self.template_sheet_combo.grid(row=4, column=1, sticky=tk.W, padx=5, pady=5)
        
        # Mapping management
        manage_mappings_btn = ttk.Button(options_frame, text="Manage Saved Mappings", 
                                       command=self.show_mappings_manager)
        manage_mappings_btn.grid(row=5, column=0, sticky=tk.W, padx=5, pady=10)
        
        advanced_btn = ttk.Button(options_frame, text="Advanced Options", 
                                command=self.show_advanced_options)
        advanced_btn.grid(row=5, column=1, sticky=tk.W, padx=5, pady=10)
    
    def create_button_section(self, parent):
        """Create the action buttons section."""
        button_frame = ttk.Frame(parent, padding="10")
        button_frame.pack(fill=tk.X, pady=10)
        
        # Process button
        process_btn = ttk.Button(button_frame, text="Process Files", 
                               command=self.process_files, width=20)
        process_btn.pack(side=tk.LEFT, padx=5)
        
        # Preview mapping button
        preview_btn = ttk.Button(button_frame, text="Preview Mapping", 
                               command=self.preview_mapping, width=20)
        preview_btn.pack(side=tk.LEFT, padx=5)
        
        # Exit button
        exit_btn = ttk.Button(button_frame, text="Exit", 
                            command=self.on_exit, width=10)
        exit_btn.pack(side=tk.RIGHT, padx=5)
    
    def create_status_section(self, parent):
        """Create the status section."""
        status_frame = ttk.Frame(parent, padding="10")
        status_frame.pack(fill=tk.X, pady=5, side=tk.BOTTOM)
        
        # Status label
        ttk.Label(status_frame, text="Status:").pack(side=tk.LEFT, padx=5)
        self.status_var = tk.StringVar(value="Ready")
        ttk.Label(status_frame, textvariable=self.status_var).pack(side=tk.LEFT, padx=5)
        
        # Progress bar
        self.progress_var = tk.DoubleVar(value=0.0)
        self.progress_bar = ttk.Progressbar(status_frame, variable=self.progress_var, 
                                          length=400, mode="determinate")
        self.progress_bar.pack(side=tk.LEFT, padx=10, fill=tk.X, expand=True)
    
    def load_settings(self):
        """Load saved settings from config."""
        # Load checkbox settings
        self.remember_dirs_var.set(self.config_mgr.get_setting("remember_directories", True))
        self.open_after_var.set(self.config_mgr.get_setting("open_after_processing", True))
        self.enable_logging_var.set(self.config_mgr.get_setting("enable_logging", False))
        self.auto_detect_var.set(self.config_mgr.get_setting("auto_detect_formats", True))
        self.use_saved_mappings_var.set(self.config_mgr.get_setting("use_saved_mappings", True))
        
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
        self.config_mgr.set_setting("open_after_processing", self.open_after_var.get())
        self.config_mgr.set_setting("enable_logging", self.enable_logging_var.get())
        self.config_mgr.set_setting("auto_detect_formats", self.auto_detect_var.get())
        self.config_mgr.set_setting("use_saved_mappings", self.use_saved_mappings_var.get())
        
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
            filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")]
        )
        
        if filename:
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
        
        if not os.path.exists(adjusted_file):
            messagebox.showerror("Error", f"Adjusted Rates file does not exist: {adjusted_file}")
            return False
            
        if not os.path.exists(template_file):
            messagebox.showerror("Error", f"Template file does not exist: {template_file}")
            return False
        
        # Check if output directory exists
        output_dir = os.path.dirname(output_file)
        if output_dir and not os.path.exists(output_dir):
            try:
                os.makedirs(output_dir)
            except Exception as e:
                messagebox.showerror("Error", f"Cannot create output directory: {str(e)}")
                return False
        
        # Check if output file is open
        try:
            if os.path.exists(output_file):
                with open(output_file, 'a'):
                    pass
        except Exception:
            messagebox.showerror("Error", f"Cannot access output file. It may be open in another program.")
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
            
            # Update progress
            self.update_status("Analyzing files...", 10)
            
            # Analyze files
            adjusted_structure = self.file_analyzer.analyze_file_structure(
                adjusted_file, adjusted_sheet)
            
            template_structure = self.file_analyzer.analyze_file_structure(
                template_file, template_sheet)
            
            # Update progress
            self.update_status("Generating mapping...", 20)
            
            # Generate mapping
            use_saved = self.use_saved_mappings_var.get()
            mapping = self.mapping_system.generate_mapping(
                adjusted_structure, use_saved_mappings=use_saved)
            
            # Check mapping confidence
            low_confidence = False
            for field, conf in self.mapping_system.mapping_confidence.items():
                if conf < 70:  # Threshold for low confidence
                    low_confidence = True
                    break
                    
            # If low confidence and auto-detect is enabled, show mapping dialog
            if low_confidence and self.auto_detect_var.get():
                self.update_status("Low confidence mapping - awaiting user input...", 30)
                
                # Put request for mapping dialog in queue
                self.msg_queue.put(("show_mapping_dialog", {
                    "adjusted_structure": adjusted_structure,
                    "mapping": mapping
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
            # Update progress
            self.update_status("Loading data...", 40)
            
            # Load data with mapping
            adjusted_df = self.data_processor.load_excel_file(adjusted_file, adjusted_sheet)
            template_df = self.data_processor.load_excel_file(template_file, template_sheet)
            
            # Update progress
            self.update_status("Transforming data...", 60)
            
            # Apply mapping
            mapped_df = self.mapping_system.apply_mapping(adjusted_df)
            
            # Transform data
            transformed_df = self.data_processor.transform_data(mapped_df)
            
            # Update progress
            self.update_status("Integrating with template...", 80)
            
            # Integrate with template
            final_df = self.data_processor.integrate_with_template(transformed_df, template_df)
            
            # Save output
            self.update_status("Saving output...", 90)
            final_df.to_excel(output_file, sheet_name=template_sheet, index=False)
            
            # Update progress
            self.update_status("Processing complete!", 100)
            
            # Optionally open the file
            if self.open_after_var.get():
                self.msg_queue.put(("open_file", {"file_path": output_file}))
            
            # Show success message
            self.msg_queue.put(("show_info", {
                "title": "Processing Complete",
                "message": f"Data has been successfully processed and saved to {output_file}"
            }))
            
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
    
    def preview_mapping(self):
        """Preview the column mapping."""
        if not self.validate_inputs():
            return
            
        adjusted_file = self.adjusted_rates_var.get()
        adjusted_sheet = self.adjusted_sheet_var.get()
        
        try:
            # Analyze file
            adjusted_structure = self.file_analyzer.analyze_file_structure(
                adjusted_file, adjusted_sheet)
                
            # Generate mapping
            use_saved = self.use_saved_mappings_var.get()
            mapping = self.mapping_system.generate_mapping(
                adjusted_structure, use_saved_mappings=use_saved)
                
            # Show mapping dialog
            self.show_mapping_dialog(adjusted_structure, mapping)
            
        except Exception as e:
            logging.error(f"Error previewing mapping: {str(e)}", exc_info=True)
            messagebox.showerror("Error", f"An error occurred while previewing mapping: {str(e)}")
    
    def show_mapping_dialog(self, adjusted_structure, mapping):
        """Show the mapping dialog to confirm column mappings."""
        # Get column names from structure
        columns = list(adjusted_structure['columns'].keys())
        
        # Create dialog
        dialog = MappingDialog(
            self, 
            source_columns=columns,
            required_fields=self.mapping_system.required_fields,
            suggested_mapping=mapping
        )
        
        # Show dialog and get result
        result_mapping, save_as_template, template_name = dialog.show()
        
        if result_mapping:
            # Update mapping
            self.mapping_system.current_mapping = result_mapping
            
            # Save mapping if requested
            if save_as_template and template_name:
                self.mapping_system.save_current_mapping(adjusted_structure, template_name)
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
    
    def show_mappings_manager(self):
        """Show the dialog to manage saved mappings."""
        # TO DO: Implement mappings manager dialog
        messagebox.showinfo("Not Implemented", "Mapping Manager will be implemented in a future version.")
    
    def show_advanced_options(self):
        """Show advanced options dialog."""
        # TO DO: Implement advanced options dialog
        messagebox.showinfo("Not Implemented", "Advanced Options will be implemented in a future version.")
    
    def update_status(self, message, progress_value):
        """Update status message and progress bar."""
        self.msg_queue.put(("update_status", {
            "message": message,
            "progress": progress_value
        }))
    
    def process_queue(self):
        """Process messages from the queue."""
        try:
            while True:
                action, params = self.msg_queue.get_nowait()
                
                if action == "update_status":
                    self.status_var.set(params["message"])
                    self.progress_var.set(params["progress"])
                    
                elif action == "show_error":
                    messagebox.showerror(params["title"], params["message"])
                    
                elif action == "show_info":
                    messagebox.showinfo(params["title"], params["message"])
                    
                elif action == "enable_controls":
                    self.enable_controls()
                    
                elif action == "open_file":
                    self.open_file(params["file_path"])
                    
                elif action == "show_mapping_dialog":
                    self.show_mapping_dialog(
                        params["adjusted_structure"],
                        params["mapping"]
                    )
                
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
        if sys.platform == 'win32':
            os.startfile(file_path)
        elif sys.platform == 'darwin':  # macOS
            os.system(f'open "{file_path}"')
        else:  # Linux
            os.system(f'xdg-open "{file_path}"')
    
    def on_exit(self):
        """Handle application exit."""
        # Save settings
        self.save_settings()
        
        logging.info("Application exiting")
        self.destroy()

def main():
    app = Application()
    app.mainloop()

if __name__ == "__main__":
    main() 