#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Mapping System module for Moxy Rates Template Transfer

This module provides functionality for mapping between different column naming
conventions and a visual interface for user-guided mapping.
"""

import os
import logging
import hashlib
import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd

class MappingSystem:
    """Handles mapping between different column naming conventions."""
    
    def __init__(self, config_manager):
        """
        Initialize the mapping system.
        
        Args:
            config_manager: MappingConfigManager instance for saving/loading mappings
        """
        self.config_manager = config_manager
        self.current_mapping = {}
        self.mapping_confidence = {}
        self.required_fields = [
            'CompanyCode', 'Term', 'Miles', 'FromMiles', 'ToMiles', 'Coverage', 'State', 'Class',
            'PlanDeduct', 'Deduct0', 'Deduct50', 'Deduct100', 'Deduct200', 'Deduct250', 'Deduct500',
            'Markup', 'New/Used', 'MaxYears', 'SurchargeCode', 'PlanCode', 'RateCardCode',
            'ClassListCode', 'MinYear', 'IncScCode', 'IncScAmt'
        ]
        
        logging.info("MappingSystem initialized")
    
    def set_required_fields(self, fields):
        """
        Set the required fields for mapping dynamically.
        
        Args:
            fields: List of field names required for mapping
        """
        if fields and isinstance(fields, list):
            self.required_fields = fields
            logging.info(f"Updated required fields: {fields}")
        else:
            logging.warning("Invalid fields parameter - must be a non-empty list")
    
    def generate_mapping(self, source_columns, use_saved_mappings=True):
        """Generate mapping between source columns and required fields."""
        mapping = {}
        self.mapping_confidence = {}
        confidence = {}
        
        # STEP 1: Ensure source_columns is properly formatted
        if isinstance(source_columns, dict):
            if 'columns' in source_columns:
                source_cols = list(source_columns['columns'].keys())
            else:
                source_cols = list(source_columns.keys())
            source_structure = source_columns
        else:
            source_cols = list(source_columns)
            source_structure = {
                'columns': {col: {} for col in source_cols},
                'column_mapping_suggestions': {}
            }
        
        # STEP 2: Use high confidence suggestions if available
        if 'column_mapping_suggestions' in source_structure:
            suggestions = source_structure['column_mapping_suggestions']
            for required_field in self.required_fields:
                field_lower = required_field.lower()
                if field_lower in suggestions:
                    suggested_column = suggestions[field_lower]['suggested_column']
                    confidence_score = suggestions[field_lower]['confidence']
                    
                    # Only use high confidence suggestions in first pass
                    if confidence_score >= 90:  # Increased threshold for stricter matching
                        mapping[required_field] = suggested_column
                        confidence[required_field] = confidence_score
        
        # STEP 3: Look for exact matches only
        for required_field in self.required_fields:
            if required_field not in mapping:
                field_lower = required_field.lower()
                
                # Look for exact matches only
                for col in source_cols:
                    col_lower = str(col).lower()
                    if col_lower == field_lower:
                        mapping[required_field] = col
                        confidence[required_field] = 100
                        break
        
        # STEP 4: Skip fuzzy matching for certain fields
        protected_fields = ['RateCardCode', 'RateCost']  # Fields that should not be auto-mapped
        
        # STEP 5: Use content-based guessing for remaining fields, except protected ones
        for required_field in self.required_fields:
            if required_field not in mapping and required_field not in protected_fields:
                field_lower = required_field.lower()
                best_match = None
                best_score = 30  # Minimum threshold
                
                for col in source_cols:
                    col_lower = str(col).lower()
                    
                    # Skip if this column is already mapped
                    if col in mapping.values():
                        continue
                    
                    # Only do exact word matching, no partial matches
                    if field_lower == col_lower:
                        best_match = col
                        best_score = 100
                        break
                    elif field_lower.replace(" ", "") == col_lower.replace(" ", ""):
                        best_match = col
                        best_score = 90
                        break
                
                if best_match:
                    mapping[required_field] = best_match
                    confidence[required_field] = best_score
        
        self.current_mapping = mapping
        self.mapping_confidence = confidence
        
        # Log mapping summary
        mapped_fields = len(mapping)
        missing_fields = len(self.required_fields) - mapped_fields
        logging.info(f"Generated mapping with {mapped_fields} fields mapped, {missing_fields} missing")
        
        return mapping
    
    def apply_mapping(self, df):
        """
        Apply the current mapping to transform a dataframe.
        
        Args:
            df: Source dataframe
            
        Returns:
            DataFrame: The original dataframe (mapping is now done in transform_data)
        """
        if not self.current_mapping:
            error_msg = "No mapping defined. Call generate_mapping first."
            logging.error(error_msg)
            raise ValueError(error_msg)
        
        # We now pass the original dataframe directly to transform_data
        # along with the mapping information
        logging.info("Mapping will be applied during transformation")
        
        # Verify all required columns exist in the source dataframe
        missing_source_cols = []
        all_source_cols = df.columns.tolist()
        
        # Check each mapping to verify source columns exist
        for standard_field, source_col in self.current_mapping.items():
            if source_col not in all_source_cols:
                missing_source_cols.append(source_col)
                logging.warning(f"Source column '{source_col}' for field '{standard_field}' not found in dataframe")
        
        # Provide detailed warnings if columns are missing
        if missing_source_cols:
            logging.warning(f"Missing source columns: {', '.join(missing_source_cols)}")
            logging.warning(f"Available columns: {', '.join(all_source_cols)}")
            
            # Try fuzzy matching to suggest alternatives
            suggestions = {}
            for missing_col in missing_source_cols:
                closest_matches = self._get_closest_matches(missing_col, all_source_cols)
                if closest_matches:
                    suggestions[missing_col] = closest_matches
                    logging.info(f"Possible alternatives for '{missing_col}': {', '.join(closest_matches)}")
        
        return df
    
    def _get_closest_matches(self, col_name, available_columns, limit=3):
        """
        Find closest matching columns using fuzzy string matching.
        
        Args:
            col_name: Column name to match
            available_columns: List of available column names
            limit: Maximum number of matches to return
            
        Returns:
            list: Best matching column names
        """
        try:
            # Import fuzzywuzzy for string matching
            from fuzzywuzzy import process
            
            # Find best matches
            matches = process.extract(col_name, available_columns, limit=limit)
            
            # Filter to only reasonable matches (score > 60)
            good_matches = [match[0] for match in matches if match[1] > 60]
            return good_matches
        except ImportError:
            # Fallback to simple substring matching if fuzzywuzzy not available
            return [col for col in available_columns if col_name.lower() in col.lower() or col.lower() in col_name.lower()]
    
    def save_current_mapping(self, source_structure, mapping_name=None):
        """
        Save the current mapping for future use.
        
        Args:
            source_structure: Source file structure for signature generation
            mapping_name: Optional name for this mapping template
        """
        if not self.current_mapping:
            error_msg = "No mapping defined to save"
            logging.error(error_msg)
            raise ValueError(error_msg)
        
        file_signature = self._generate_file_signature(source_structure)
        self.config_manager.save_mapping(
            file_signature, 
            self.current_mapping,
            mapping_name=mapping_name
        )
        
        logging.info(f"Saved mapping with signature {file_signature}" + 
                    (f" and name '{mapping_name}'" if mapping_name else ""))
    
    def _generate_file_signature(self, structure):
        """
        Generate a unique signature for a file structure.
        
        Args:
            structure: File structure analysis result
            
        Returns:
            str: MD5 hash signature of the file structure
        """
        # Create a signature based on column names, order, and data types
        columns = sorted(list(structure.get('columns', {}).keys()))
        data_types = [structure.get('columns', {}).get(col, {}).get('data_type', 'unknown') 
                     for col in columns]
        
        signature_data = str(columns).encode() + str(data_types).encode()
        signature = hashlib.md5(signature_data).hexdigest()
        
        logging.debug(f"Generated file signature: {signature}")
        return signature


class MappingDialog:
    """Visual interface for manually mapping columns between files."""
    
    def __init__(self, parent, source_columns, required_fields, suggested_mapping=None):
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Column Mapping - Moxy Rates Template Transfer")
        self.dialog.geometry("900x700")  # Increased height to ensure all content is visible
        self.dialog.minsize(800, 600)
        
        # Make dialog modal
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        # Store parent reference and ensure it stays dark
        self.parent = parent
        self._original_parent_bg = parent.cget('background')  # Store original parent background
        
        # Configure dialog colors
        self.DARK_BG = "#1E1E1E"
        self.DARKER_BG = "#171717"
        self.TEXT_COLOR = "#FFFFFF"
        self.ACCENT_COLOR = "#8DC63F"
        self.BUTTON_HOVER_BG = "#2B5BA0"

        # Configure dialog background first
        self.dialog.configure(background=self.DARK_BG)
        
        # Configure system-level styles for this dialog's widgets
        self.dialog.option_add('*TCombobox*Listbox.background', self.DARKER_BG)
        self.dialog.option_add('*TCombobox*Listbox.foreground', self.TEXT_COLOR)
        self.dialog.option_add('*TCombobox*Listbox.selectBackground', self.ACCENT_COLOR)
        self.dialog.option_add('*TCombobox*Listbox.selectForeground', self.TEXT_COLOR)
        self.dialog.option_add('*Entry.background', self.DARKER_BG)
        self.dialog.option_add('*Entry.foreground', self.TEXT_COLOR)
        self.dialog.option_add('*Entry.insertBackground', self.TEXT_COLOR)
        self.dialog.option_add('*Listbox.background', self.DARKER_BG)
        self.dialog.option_add('*Listbox.foreground', self.TEXT_COLOR)
        self.dialog.option_add('*Listbox.selectBackground', self.ACCENT_COLOR)
        self.dialog.option_add('*Listbox.selectForeground', self.TEXT_COLOR)
        
        # Store parameters
        self.source_columns = list(source_columns) if not isinstance(source_columns, list) else source_columns
        self.required_fields = list(required_fields) if not isinstance(required_fields, list) else required_fields
        self.mapping = suggested_mapping or {}
        self.result_mapping = None
        self.save_as_template = False
        self.template_name = ""
        
        # Create main container with padding
        self.main_container = ttk.Frame(self.dialog, padding="10", style="Dialog.TFrame")
        self.main_container.pack(fill=tk.BOTH, expand=True)
        
        # Configure styles before creating UI components
        self._configure_styles()
        
        # Create UI components
        self._create_ui()
        
        # Bind dialog events for theme maintenance
        self.dialog.bind("<Destroy>", self._on_dialog_close)
        self.dialog.bind("<Map>", self._on_dialog_map)
        self.dialog.bind("<Unmap>", self._on_dialog_unmap)
        
        # Set up parent theme preservation
        self._ensure_parent_stays_dark()
    
    def _configure_styles(self):
        """Configure the styles for the dialog."""
        style = ttk.Style()
        current_theme = style.theme_use()  # Store current theme
        
        # Configure dialog-specific styles without changing the global theme
        style.configure("Dialog.TCombobox",
                       background=self.DARKER_BG,
                       foreground=self.TEXT_COLOR,
                       fieldbackground=self.DARKER_BG,
                       selectbackground=self.ACCENT_COLOR,
                       selectforeground=self.TEXT_COLOR,
                       borderwidth=1,
                       padding=5,
                       arrowcolor=self.TEXT_COLOR)
        
        style.map("Dialog.TCombobox",
                 fieldbackground=[("readonly", self.DARKER_BG),
                                ("disabled", self.DARK_BG),
                                ("active", self.DARKER_BG),
                                ("focus", self.DARKER_BG)],
                 selectbackground=[("readonly", self.ACCENT_COLOR)],
                 selectforeground=[("readonly", self.TEXT_COLOR)],
                 background=[("readonly", self.DARKER_BG),
                           ("disabled", self.DARK_BG),
                           ("active", self.DARKER_BG),
                           ("focus", self.DARKER_BG)],
                 foreground=[("readonly", self.TEXT_COLOR),
                           ("disabled", "#666666"),
                           ("active", self.TEXT_COLOR),
                           ("focus", self.TEXT_COLOR)])

        # Configure button style
        style.configure("Dialog.TButton",
                       background=self.DARKER_BG,
                       foreground=self.TEXT_COLOR,
                       bordercolor=self.DARK_BG,
                       darkcolor=self.DARK_BG,
                       lightcolor=self.DARK_BG,
                       relief="flat",
                       font=("Segoe UI", 9))
        
        # Configure dialog style
        style.configure("Dialog.TFrame", background=self.DARK_BG)
        style.configure("Dialog.TLabel",
                       background=self.DARK_BG,
                       foreground=self.TEXT_COLOR,
                       font=("Segoe UI", 9))
        style.configure("DialogHeader.TLabel",
                       background=self.DARK_BG,
                       foreground=self.TEXT_COLOR,
                       font=("Segoe UI", 14, "bold"))
        style.configure("DialogNote.TLabel",
                       background=self.DARKER_BG,
                       foreground=self.TEXT_COLOR,
                       font=("Segoe UI", 9))
        
        # Configure additional dialog-specific styles
        style.configure("Dialog.TLabelframe",
                       background=self.DARKER_BG,
                       bordercolor=self.DARK_BG,
                       darkcolor=self.DARK_BG,
                       lightcolor=self.DARK_BG)
        style.configure("Dialog.TLabelframe.Label",
                       background=self.DARK_BG,
                       foreground=self.TEXT_COLOR,
                       font=("Segoe UI", 10, "bold"))
        
        # Entry fields in dialog
        style.configure("Dialog.TEntry",
                       background=self.DARKER_BG,
                       fieldbackground=self.DARKER_BG,
                       foreground=self.TEXT_COLOR,
                       insertcolor=self.TEXT_COLOR,
                       borderwidth=1,
                       relief="solid",
                       padding=8)
        style.map("Dialog.TEntry",
                 fieldbackground=[("disabled", self.DARK_BG),
                                ("readonly", self.DARKER_BG)],
                 background=[("disabled", self.DARK_BG),
                           ("readonly", self.DARKER_BG)])
        
        # Configure checkbutton style
        style.configure("Dialog.TCheckbutton",
                       background=self.DARKER_BG,
                       foreground=self.TEXT_COLOR,
                       selectcolor=self.DARKER_BG)
        style.map("Dialog.TCheckbutton",
                 background=[("active", self.DARKER_BG)],
                 foreground=[("active", self.TEXT_COLOR)])

        # Update labelframe style to ensure dark background
        style.configure("Dialog.TLabelframe",
                       background=self.DARKER_BG,
                       darkcolor=self.DARKER_BG,
                       lightcolor=self.DARKER_BG)
        style.configure("Dialog.TLabelframe.Label",
                       background=self.DARKER_BG,
                       foreground=self.TEXT_COLOR,
                       font=("Segoe UI", 10, "bold"))
        
    def _create_ui(self):
        """Create the UI components for mapping."""
        # Title and instructions at the top
        title_frame = ttk.Frame(self.main_container, style="Dialog.TFrame")
        title_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(title_frame, 
                 text="Map Source Columns to Required Fields",
                 style="DialogHeader.TLabel").pack(anchor=tk.W)
        
        ttk.Label(title_frame,
                 text="Select the source column from your data file that corresponds to each required field in the template.",
                 style="Dialog.TLabel",
                 wraplength=850).pack(anchor=tk.W, pady=(5, 0))
        
        # Important note section with darker background
        note_frame = ttk.LabelFrame(self.main_container, 
                                  text="Important Note",
                                  padding="10",
                                  style="Dialog.TLabelframe")
        note_frame.pack(fill=tk.X, pady=10)
        
        note_text = ("Deductible and RateCost columns are handled automatically to populate "
                    "deductible-specific columns (e.g., Deduct0, Deduct50) in the final output. "
                    "Make sure these fields are correctly mapped.")
        ttk.Label(note_frame, 
                 text=note_text,
                 style="DialogNote.TLabel",
                 wraplength=850).pack(anchor=tk.W)
        
        # Mapping section
        mapping_frame = ttk.LabelFrame(self.main_container,
                                     text="Column Mapping",
                                     padding="10",
                                     style="Dialog.TLabelframe")
        mapping_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # Headers with improved contrast
        headers_frame = ttk.Frame(mapping_frame, style="Dialog.TFrame")
        headers_frame.pack(fill=tk.X)
        
        ttk.Label(headers_frame,
                 text="Required Field",
                 width=20,
                 style="Dialog.TLabel",
                 font=("Segoe UI", 10, "bold")).pack(side=tk.LEFT, padx=5)
        ttk.Label(headers_frame,
                 text="Source Column",
                 width=40,
                 style="Dialog.TLabel",
                 font=("Segoe UI", 10, "bold")).pack(side=tk.LEFT, padx=5)
        ttk.Label(headers_frame,
                 text="Status",
                 width=20,
                 style="Dialog.TLabel",
                 font=("Segoe UI", 10, "bold")).pack(side=tk.LEFT, padx=5)
        
        # Scrollable frame for mappings
        canvas_frame = tk.Frame(mapping_frame, bg=self.DARKER_BG)  # Changed to tk.Frame
        canvas_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # Create canvas and scrollbar
        self.canvas = tk.Canvas(canvas_frame,
                              bg=self.DARKER_BG,  # Use bg instead of background
                              highlightthickness=0)
        scrollbar = ttk.Scrollbar(canvas_frame, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = tk.Frame(self.canvas, bg=self.DARKER_BG)  # Changed to tk.Frame
        
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=scrollbar.set)
        
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Add mapping rows
        self.mapping_vars = {}
        self.preview_labels = {}
        
        for field in self.required_fields:
            row_frame = tk.Frame(self.scrollable_frame, bg=self.DARKER_BG)
            row_frame.pack(fill=tk.X, pady=2)
            
            # Field name
            tk.Label(row_frame,
                    text=field,
                    width=20,
                    bg=self.DARKER_BG,
                    fg=self.TEXT_COLOR,
                    font=("Segoe UI", 9)).pack(side=tk.LEFT, padx=5)
            
            # Dropdown with explicit style
            var = tk.StringVar()
            self.mapping_vars[field] = var
            if field in self.mapping:
                var.set(self.mapping[field])
            
            dropdown = ttk.Combobox(row_frame,
                                  textvariable=var, 
                                  values=[""] + self.source_columns,
                                  width=40,
                                  style="Dialog.TCombobox",  # Use dialog-specific style
                                  state="readonly")
            dropdown.pack(side=tk.LEFT, padx=5)
            
            # Configure dropdown colors explicitly
            dropdown.configure(background=self.DARKER_BG,
                            foreground=self.TEXT_COLOR)
            
            # Status label
            preview_label = tk.Label(row_frame,
                                   width=20,
                                   bg=self.DARKER_BG,
                                   fg=self.TEXT_COLOR,
                                   font=("Segoe UI", 9))
            preview_label.pack(side=tk.LEFT, padx=5)
            self.preview_labels[field] = preview_label
            
            # Update preview when selection changes
            var.trace_add("write", lambda name, index, mode, f=field: self._update_preview(f))
            self._update_preview(field)
        
        # Template saving section
        save_frame = ttk.LabelFrame(self.main_container,
                                  text="Save Mapping Template",
                                  padding="10",
                                  style="Dialog.TLabelframe")  # Use our dark style
        save_frame.pack(fill=tk.X, pady=10)
        
        # Create a frame for the checkbox with dark background
        checkbox_frame = ttk.Frame(save_frame, style="Dialog.TFrame")
        checkbox_frame.pack(fill=tk.X)
        
        self.save_template_var = tk.BooleanVar(value=False)
        check = ttk.Checkbutton(checkbox_frame, 
                              text="Save this mapping as a template for future use",
                              variable=self.save_template_var,
                              style="Dialog.TCheckbutton")  # Use custom checkbutton style
        check.pack(anchor=tk.W)
        
        template_name_frame = ttk.Frame(save_frame, style="Dialog.TFrame")
        template_name_frame.pack(fill=tk.X, pady=(5, 0))
        
        ttk.Label(template_name_frame,
                 text="Template name:",
                 style="Dialog.TLabel").pack(side=tk.LEFT, padx=(0, 5))
        
        self.template_name_var = tk.StringVar()
        template_entry = tk.Entry(template_name_frame,
                                textvariable=self.template_name_var,
                                width=40,
                                bg=self.DARKER_BG,
                                fg=self.TEXT_COLOR,
                                insertbackground=self.TEXT_COLOR,
                                relief='solid',
                                bd=1)
        template_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # Buttons at the bottom
        button_frame = ttk.Frame(self.main_container, style="Dialog.TFrame")
        button_frame.pack(fill=tk.X, pady=10)
        
        # Auto-map button with Moxy blue styling
        auto_map_btn = tk.Button(button_frame,
                                text="Auto-Map Remaining",
                                bg=self.BUTTON_HOVER_BG,
                                fg='#FFFFFF',
                                font=("Segoe UI", 9, "bold"),
                                relief='flat',
                                activebackground=self.BUTTON_HOVER_BG,
                                activeforeground='#FFFFFF',
                                width=20,
                                command=self._auto_map_remaining)
        auto_map_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        # Right side buttons
        right_buttons = ttk.Frame(button_frame, style="Dialog.TFrame")
        right_buttons.pack(side=tk.RIGHT)
        
        # Cancel button with Moxy blue styling
        cancel_btn = tk.Button(right_buttons,
                              text="Cancel",
                              bg=self.BUTTON_HOVER_BG,
                              fg='#FFFFFF',
                              font=("Segoe UI", 9, "bold"),
                              relief='flat',
                              activebackground=self.BUTTON_HOVER_BG,
                              activeforeground='#FFFFFF',
                              width=15,
                              command=self.dialog.destroy)
        cancel_btn.pack(side=tk.RIGHT, padx=(5, 0))
        
        # Apply button with Moxy green styling
        apply_btn = tk.Button(right_buttons,
                             text="Apply Mapping",
                             bg=self.ACCENT_COLOR,
                             fg='#000000',
                             font=("Segoe UI", 9, "bold"),
                             relief='flat',
                             activebackground="#9ED84F",
                             activeforeground='#000000',
                             width=15,
                             command=self._apply_mapping)
        apply_btn.pack(side=tk.RIGHT, padx=5)
        
        # Configure system-level styles for combobox popdown
        self.dialog.option_add('*TCombobox*Listbox.background', self.DARKER_BG)
        self.dialog.option_add('*TCombobox*Listbox.foreground', self.TEXT_COLOR)
        self.dialog.option_add('*TCombobox*Listbox.selectBackground', self.ACCENT_COLOR)
        self.dialog.option_add('*TCombobox*Listbox.selectForeground', self.TEXT_COLOR)
        self.dialog.option_add('*TCombobox*Listbox.font', ("Segoe UI", 9))
        
        # Add scrolling support
        self._add_scrolling_support()
    
    def _update_preview(self, field):
        """
        Update the preview for a field based on the current selection.
        
        Args:
            field: Field name to update preview for
        """
        selected_column = self.mapping_vars[field].get()
        if not selected_column:
            preview_text = "(No column selected)"
            self.preview_labels[field].config(text=preview_text,
                                           foreground="#FF4444")  # Bright red for better visibility
        else:
            preview_text = f"Selected: {selected_column}"
            self.preview_labels[field].config(text=preview_text,
                                           foreground="#8DC63F")  # Moxy green for success
    
    def _auto_map_remaining(self):
        """Attempt to automatically map remaining unmapped fields."""
        # Count unmapped fields before
        unmapped_before = sum(1 for field in self.required_fields if not self.mapping_vars[field].get())
        
        # Get list of already mapped source columns
        used_columns = [var.get() for var in self.mapping_vars.values() if var.get()]
        
        # Get list of available columns (not already mapped)
        available_columns = [col for col in self.source_columns if col not in used_columns]
        
        # Try to map remaining fields using heuristics
        for field in self.required_fields:
            if not self.mapping_vars[field].get() and available_columns:
                # Use fuzzy matching to find the best match
                best_match = None
                best_score = 0
                
                field_lower = field.lower()
                for col in available_columns:
                    col_lower = col.lower()
                    
                    # Exact match is best
                    if col_lower == field_lower:
                        best_match = col
                        break
                    
                    # Contains field name
                    if field_lower in col_lower:
                        best_match = col
                        break
                    
                    # Contains word parts
                    for word in self._split_camel_case(field_lower):
                        if word in col_lower and len(word) > 2:  # Avoid short words
                            best_match = col
                            break
                
                if best_match:
                    self.mapping_vars[field].set(best_match)
                    available_columns.remove(best_match)
        
        # Count unmapped fields after
        unmapped_after = sum(1 for field in self.required_fields if not self.mapping_vars[field].get())
        mapped_count = unmapped_before - unmapped_after
        
        if mapped_count > 0:
            messagebox.showinfo("Auto-Mapping", f"Successfully auto-mapped {mapped_count} fields.")
        else:
            messagebox.showinfo("Auto-Mapping", "No additional fields could be auto-mapped.")
    
    def _split_camel_case(self, name):
        """
        Split a camel case name into individual words.
        
        Args:
            name: String to split
            
        Returns:
            list: Individual words
        """
        result = []
        current_word = ""
        
        for char in name:
            if char.isupper() and current_word:
                result.append(current_word.lower())
                current_word = char
            else:
                current_word += char
        
        if current_word:
            result.append(current_word.lower())
            
        # Also add word splits by underscore and space
        for word in " ".join(result).replace("_", " ").split():
            if word not in result:
                result.append(word)
                
        return result
    
    def _apply_mapping(self):
        """Apply the mapping and close the dialog."""
        # Get the final mapping from the UI
        self.result_mapping = {
            field: var.get() for field, var in self.mapping_vars.items()
            if var.get()  # Only include mapped fields
        }
        
        # Check if any required fields are unmapped
        unmapped = [field for field in self.required_fields if field not in self.result_mapping]
        if unmapped:
            message = f"The following required fields are not mapped: {', '.join(unmapped)}\n\nDo you want to continue?"
            if not messagebox.askyesno("Warning", message):
                return
        
        # Check if we should save as template
        self.save_as_template = self.save_template_var.get()
        self.template_name = self.template_name_var.get() if self.save_as_template else ""
        
        if self.save_as_template and not self.template_name:
            messagebox.showerror("Error", "Please enter a name for the mapping template.")
            return
        
        logging.info(f"Applied mapping with {len(self.result_mapping)} fields mapped")
        
        # Close the dialog
        self.dialog.destroy()
    
    def show(self):
        """
        Show the dialog and return the mapping when closed.
        
        Returns:
            tuple: (mapping, save_as_template, template_name)
        """
        # Wait for the dialog to close
        self.dialog.wait_window()
        
        return self.result_mapping, self.save_as_template, self.template_name
    
    def _add_scrolling_support(self):
        """Add support for touchpad scrolling and mouse wheel"""
        def _on_mousewheel(event):
            # Get direction of scroll
            if event.num == 4 or event.delta > 0:
                self.canvas.yview_scroll(-1, "units")
            elif event.num == 5 or event.delta < 0:
                self.canvas.yview_scroll(1, "units")
            return "break"  # Prevent event propagation
            
        def _on_touchpad(event):
            # Touchpad scrolling for Windows (uses event.delta)
            if event.delta:
                self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
            return "break"
        
        # Bind mouse wheel event for Linux (event.num) and Windows/Mac (event.delta)
        self.canvas.bind_all("<MouseWheel>", _on_touchpad)  # Windows/Mac
        self.canvas.bind_all("<Button-4>", _on_mousewheel)  # Linux scroll up
        self.canvas.bind_all("<Button-5>", _on_mousewheel)  # Linux scroll down
        
        # Also bind for touchpad on Windows (may use different event)
        self.canvas.bind_all("<MouseWheel>", _on_touchpad)
        
        # Unbind these events when the dialog is destroyed
        self.dialog.bind("<Destroy>", self._remove_scroll_bindings)
        
    def _remove_scroll_bindings(self, event):
        """Remove all scroll bindings when dialog is closed to avoid interference"""
        try:
            self.canvas.unbind_all("<MouseWheel>")
            self.canvas.unbind_all("<Button-4>")
            self.canvas.unbind_all("<Button-5>")
            logging.info("Removed scroll bindings from mapping dialog")
        except Exception as e:
            logging.error(f"Error removing scroll bindings: {str(e)}")
        
        # Check if the event's widget is this dialog
        if event.widget == self.dialog:
            logging.info("Mapping dialog destroyed, bindings removed")

    def _ensure_parent_stays_dark(self):
        """Ensure the parent window maintains its dark theme"""
        if self.parent:
            # Set up observer to maintain dark theme
            def _observe_parent_bg(*args):
                if self.parent.winfo_exists():
                    current_bg = self.parent.cget('background')
                    if current_bg != self._original_parent_bg:
                        self.parent.configure(background=self._original_parent_bg)
            
            # Create a StringVar to track background changes
            self._parent_bg_var = tk.StringVar(value=self._original_parent_bg)
            self._parent_bg_var.trace_add("write", _observe_parent_bg)
            
            # Bind to parent window events that might affect theming
            self.parent.bind("<Map>", lambda e: _observe_parent_bg())
            self.parent.bind("<Expose>", lambda e: _observe_parent_bg())

    def _on_dialog_close(self, event):
        """Handle dialog close event to maintain parent window theme"""
        try:
            if event.widget == self.dialog:
                # Remove scroll bindings
                self._remove_scroll_bindings(event)
                
                # Restore parent window theme
                if self.parent and self.parent.winfo_exists():
                    self.parent.configure(background=self._original_parent_bg)
                    
                # Clean up parent window bindings
                if hasattr(self, '_parent_bg_var'):
                    self.parent.unbind("<Map>")
                    self.parent.unbind("<Expose>")
        except Exception as e:
            logging.error(f"Error in dialog close handler: {str(e)}")

    def _on_dialog_map(self, event):
        """Handle dialog map event (when dialog becomes visible)"""
        try:
            # Ensure dialog stays dark when shown
            self.dialog.configure(background=self.DARK_BG)
            # Ensure parent maintains its original theme
            if self.parent and self.parent.winfo_exists():
                self.parent.configure(background=self._original_parent_bg)
        except Exception as e:
            logging.error(f"Error in dialog map handler: {str(e)}")

    def _on_dialog_unmap(self, event):
        """Handle dialog unmap event (when dialog is hidden)"""
        try:
            # Ensure parent maintains its original theme when dialog is hidden
            if self.parent and self.parent.winfo_exists():
                self.parent.configure(background=self._original_parent_bg)
        except Exception as e:
            logging.error(f"Error in dialog unmap handler: {str(e)}") 