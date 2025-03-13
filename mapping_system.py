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
            'Coverage', 'Term', 'Miles', 'FromMiles', 'ToMiles', 
            'MinYear', 'MaxYears', 'Class', 'RateCost', 'Deductible'
        ]
        
        logging.info("MappingSystem initialized")
    
    def generate_mapping(self, source_structure, target_structure=None, use_saved_mappings=True):
        """
        Generate mapping between source and standardized columns.
        
        Args:
            source_structure: Source file structure analysis
            target_structure: Optional target structure for direct mapping
            use_saved_mappings: Whether to use saved mappings if available
            
        Returns:
            dict: Mapping between source and standard columns
        """
        logging.info("Generating mapping for file structure")
        
        # If we have a saved mapping for this file pattern and should use it
        if use_saved_mappings:
            file_signature = self._generate_file_signature(source_structure)
            saved_mapping = self.config_manager.get_saved_mapping(file_signature)
            
            if saved_mapping:
                logging.info(f"Using saved mapping with signature: {file_signature}")
                self.current_mapping = saved_mapping
                self.mapping_confidence = {field: 100 for field in saved_mapping}
                return saved_mapping
        
        # Otherwise, generate mapping based on column analysis
        mapping = {}
        confidence = {}
        
        # Use suggested mappings from file analysis
        suggestions = source_structure.get('column_mapping_suggestions', {})
        
        for required_field in self.required_fields:
            field_lower = required_field.lower()
            if field_lower in suggestions:
                suggested_column = suggestions[field_lower]['suggested_column']
                confidence_score = suggestions[field_lower]['confidence']
                
                mapping[required_field] = suggested_column
                confidence[required_field] = confidence_score
                
                logging.debug(f"Mapped {required_field} -> {suggested_column} (confidence: {confidence_score})")
            else:
                # Fall back to direct name matching
                for col in source_structure.get('columns', {}):
                    if col.lower() == field_lower:
                        mapping[required_field] = col
                        confidence[required_field] = 100
                        logging.debug(f"Direct match {required_field} -> {col}")
                        break
        
        self.current_mapping = mapping
        self.mapping_confidence = confidence
        
        # Log mapping summary
        mapped_fields = len(mapping)
        missing_fields = self.required_fields.count - mapped_fields if hasattr(self.required_fields, 'count') else len(self.required_fields) - mapped_fields
        logging.info(f"Generated mapping with {mapped_fields} fields mapped, {missing_fields} missing")
        
        return mapping
    
    def apply_mapping(self, df):
        """
        Apply the current mapping to transform a dataframe.
        
        Args:
            df: Source dataframe
            
        Returns:
            DataFrame: Transformed dataframe with standardized columns
        """
        if not self.current_mapping:
            error_msg = "No mapping defined. Call generate_mapping first."
            logging.error(error_msg)
            raise ValueError(error_msg)
        
        # Create new dataframe with mapped columns
        mapped_df = pd.DataFrame()
        
        for standard_col, source_col in self.current_mapping.items():
            if source_col in df.columns:
                mapped_df[standard_col] = df[source_col]
                logging.debug(f"Applied mapping: {source_col} -> {standard_col}")
            else:
                logging.warning(f"Column {source_col} not found in dataframe for mapping to {standard_col}")
        
        # Check for missing required columns
        missing = [col for col in self.required_fields if col not in mapped_df.columns]
        if missing:
            error_msg = f"Missing required columns after mapping: {', '.join(missing)}"
            logging.error(error_msg)
            raise ValueError(error_msg)
        
        logging.info(f"Successfully mapped dataframe with {len(mapped_df.columns)} columns")
        return mapped_df
    
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
        """
        Initialize the mapping dialog.
        
        Args:
            parent: Parent window
            source_columns: List of column names from source file
            required_fields: List of required standard fields
            suggested_mapping: Optional dictionary of suggested mappings
        """
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Column Mapping")
        self.dialog.geometry("800x600")
        self.dialog.minsize(800, 600)
        
        # Make dialog modal
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        self.source_columns = source_columns
        self.required_fields = required_fields
        self.mapping = suggested_mapping or {}
        self.result_mapping = None
        self.save_as_template = False
        self.template_name = ""
        
        logging.info("MappingDialog initialized")
        self._create_ui()
    
    def _create_ui(self):
        """Create the UI components for mapping."""
        # Header frame
        header_frame = ttk.Frame(self.dialog, padding="10")
        header_frame.pack(fill=tk.X)
        
        ttk.Label(header_frame, 
                text="Map Source Columns to Required Fields", 
                font=("Arial", 14, "bold")).pack(side=tk.LEFT)
        
        # Instructions
        instr_frame = ttk.Frame(self.dialog, padding="10")
        instr_frame.pack(fill=tk.X)
        
        ttk.Label(instr_frame, 
                text="Select the source column that corresponds to each required field.",
                font=("Arial", 10)).pack(anchor=tk.W)
        
        # Mapping area with scrolling
        mapping_frame = ttk.Frame(self.dialog, padding="10")
        mapping_frame.pack(fill=tk.BOTH, expand=True)
        
        # Canvas with scrollbar for mapping entries
        canvas = tk.Canvas(mapping_frame)
        scrollbar = ttk.Scrollbar(mapping_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Create a dropdown for each required field
        self.mapping_vars = {}
        self.preview_labels = {}
        
        # Add headers
        ttk.Label(scrollable_frame, text="Required Field", font=("Arial", 10, "bold"), 
                width=20).grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        ttk.Label(scrollable_frame, text="Source Column", font=("Arial", 10, "bold"), 
                width=40).grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
        ttk.Label(scrollable_frame, text="Preview", font=("Arial", 10, "bold"), 
                width=20).grid(row=0, column=2, padx=5, pady=5, sticky=tk.W)
        
        # Add mapping rows
        for i, field in enumerate(self.required_fields, 1):
            # Required field name
            ttk.Label(scrollable_frame, text=field, width=20).grid(
                row=i, column=0, padx=5, pady=5, sticky=tk.W)
            
            # Dropdown for source column
            var = tk.StringVar()
            self.mapping_vars[field] = var
            
            # Set initial value if in suggested mapping
            if field in self.mapping:
                var.set(self.mapping[field])
            
            # Create dropdown with source columns
            dropdown = ttk.Combobox(scrollable_frame, textvariable=var, 
                                   values=[""] + self.source_columns, width=40)
            dropdown.grid(row=i, column=1, padx=5, pady=5, sticky=tk.W)
            
            # Add binding to update preview when selection changes
            var.trace_add("write", lambda name, index, mode, field=field: self._update_preview(field))
            
            # Preview label
            preview_label = ttk.Label(scrollable_frame, text="", width=20)
            preview_label.grid(row=i, column=2, padx=5, pady=5, sticky=tk.W)
            self.preview_labels[field] = preview_label
            
            # Initial preview update
            self._update_preview(field)
        
        # Bottom button frame
        button_frame = ttk.Frame(self.dialog, padding="10")
        button_frame.pack(fill=tk.X, side=tk.BOTTOM)
        
        # Auto-map button
        auto_btn = ttk.Button(button_frame, text="Auto-Map Remaining", 
                             command=self._auto_map_remaining)
        auto_btn.pack(side=tk.LEFT, padx=5)
        
        # Save as template checkbox
        self.save_template_var = tk.BooleanVar(value=False)
        save_cb = ttk.Checkbutton(button_frame, text="Save as mapping template", 
                                 variable=self.save_template_var)
        save_cb.pack(side=tk.LEFT, padx=20)
        
        # Template name entry
        self.template_name_var = tk.StringVar()
        ttk.Label(button_frame, text="Template name:").pack(side=tk.LEFT, padx=5)
        name_entry = ttk.Entry(button_frame, textvariable=self.template_name_var, width=20)
        name_entry.pack(side=tk.LEFT, padx=5)
        
        # Cancel/Apply buttons
        cancel_btn = ttk.Button(button_frame, text="Cancel", command=self.dialog.destroy)
        cancel_btn.pack(side=tk.RIGHT, padx=5)
        
        apply_btn = ttk.Button(button_frame, text="Apply Mapping", command=self._apply_mapping)
        apply_btn.pack(side=tk.RIGHT, padx=5)
    
    def _update_preview(self, field):
        """
        Update the preview for a field based on the current selection.
        
        Args:
            field: Field name to update preview for
        """
        selected_column = self.mapping_vars[field].get()
        preview_text = "(No column selected)" if not selected_column else f"Selected: {selected_column}"
        self.preview_labels[field].config(text=preview_text)
    
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