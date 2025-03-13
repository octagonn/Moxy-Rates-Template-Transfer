#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Test script for Moxy Rates Template Transfer application

This script tests the core functionality of the application modules without
launching the full GUI interface.
"""

import os
import logging
import pandas as pd
import tkinter as tk
from tkinter import messagebox

# Import application modules
from file_analyzer import FileAnalyzer
from mapping_system import MappingSystem
from config_manager import ConfigManager, MappingConfigManager
from data_processor import DataProcessor

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)

def create_test_data():
    """Create test Excel files for testing."""
    print("Creating test data files...")
    
    # Create a test directory
    test_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "test_data")
    os.makedirs(test_dir, exist_ok=True)
    
    # Create a sample adjusted rates file
    adjusted_rates_data = {
        'Coverage': ['Basic', 'Premium', 'Premium', 'Basic', 'Basic'],
        'Term': [12, 24, 36, 48, 60],
        'Miles': [12000, 24000, 36000, 48000, 60000],
        'FromMiles': [0, 0, 0, 0, 0],
        'ToMiles': [12000, 24000, 36000, 48000, 60000],
        'MinYear': [2010, 2012, 2014, 2016, 2018],
        'MaxYears': [5, 7, 9, 11, 13],
        'Class': ['A', 'B', 'C', 'A', 'B'],
        'RateCost': [100.50, 200.75, 300.25, 400.00, 500.50],
        'Deductible': [0, 50, 100, 150, 200]
    }
    
    adjusted_df = pd.DataFrame(adjusted_rates_data)
    adjusted_file = os.path.join(test_dir, "test_adjusted_rates.xlsx")
    adjusted_df.to_excel(adjusted_file, sheet_name="Dealer Cost Rates", index=False)
    print(f"Created adjusted rates test file: {adjusted_file}")
    
    # Create a sample template file
    template_data = {
        'Coverage': ['Template'],
        'Term': [0],
        'Miles': [0],
        'FromMiles': [0],
        'ToMiles': [0],
        'MinYear': [0],
        'MaxYears': [0],
        'Class': ['X'],
        'Deduct0': [0],
        'Deduct50': [0],
        'Deduct100': [0],
        'Deduct150': [0],
        'Deduct200': [0],
        'PlanDeduct': [0]
    }
    
    template_df = pd.DataFrame(template_data)
    template_file = os.path.join(test_dir, "test_template.xlsx")
    template_df.to_excel(template_file, sheet_name="Sheet1", index=False)
    print(f"Created template test file: {template_file}")
    
    return adjusted_file, template_file

def test_file_analyzer(file_path):
    """Test the FileAnalyzer class."""
    print("\nTesting FileAnalyzer...")
    
    analyzer = FileAnalyzer()
    
    # Test sheet names
    sheets = analyzer.get_sheet_names(file_path)
    print(f"Detected sheets: {sheets}")
    
    # Test file structure analysis
    structure = analyzer.analyze_file_structure(file_path, sheets[0])
    print(f"Analyzed file structure. Found {len(structure.get('columns', {}))} columns.")
    
    # Print column mapping suggestions
    suggestions = structure.get('column_mapping_suggestions', {})
    print("\nColumn mapping suggestions:")
    for field, info in suggestions.items():
        print(f"  {field} -> {info['suggested_column']} (confidence: {info['confidence']})")
    
    return structure

def test_mapping_system(source_structure):
    """Test the MappingSystem class."""
    print("\nTesting MappingSystem...")
    
    # Initialize config manager and mapping system
    config_mgr = MappingConfigManager()
    mapping_system = MappingSystem(config_mgr)
    
    # Generate mapping
    mapping = mapping_system.generate_mapping(source_structure)
    
    print("Generated mapping:")
    for field, column in mapping.items():
        confidence = mapping_system.mapping_confidence.get(field, 0)
        print(f"  {field} -> {column} (confidence: {confidence})")
    
    return mapping_system, mapping

def test_data_processor(adjusted_file, template_file, mapping):
    """Test the DataProcessor class."""
    print("\nTesting DataProcessor...")
    
    processor = DataProcessor()
    
    # Load files
    adjusted_df = processor.load_excel_file(adjusted_file, "Dealer Cost Rates")
    template_df = processor.load_excel_file(template_file, "Sheet1")
    
    print(f"Loaded adjusted rates: {len(adjusted_df)} rows, {len(adjusted_df.columns)} columns")
    print(f"Loaded template: {len(template_df)} rows, {len(template_df.columns)} columns")
    
    # Create a mock mapped dataframe to simulate the mapping result
    # (In real app, this would be done by the mapping system)
    mapped_df = adjusted_df.copy()
    
    # Transform data
    transformed_df = processor.transform_data(mapped_df)
    print(f"Transformed data: {len(transformed_df)} rows, {len(transformed_df.columns)} columns")
    
    # Deductible columns
    deduct_cols = [col for col in transformed_df.columns if 'Deduct' in str(col)]
    print(f"Created deductible columns: {deduct_cols}")
    
    # Integrate with template
    final_df = processor.integrate_with_template(transformed_df, template_df)
    print(f"Final integrated data: {len(final_df)} rows, {len(final_df.columns)} columns")
    
    # Save the output to a test file
    output_file = adjusted_file.replace("adjusted_rates.xlsx", "output.xlsx")
    final_df.to_excel(output_file, sheet_name="Processed", index=False)
    print(f"Saved output to: {output_file}")
    
    return output_file

def main():
    """Run all tests."""
    print("Running Moxy Rates Template Transfer tests...\n")
    
    try:
        # Create test data
        adjusted_file, template_file = create_test_data()
        
        # Test file analyzer
        structure = test_file_analyzer(adjusted_file)
        
        # Test mapping system
        mapping_system, mapping = test_mapping_system(structure)
        
        # Test data processor
        output_file = test_data_processor(adjusted_file, template_file, mapping)
        
        print("\nAll tests completed successfully!")
        print(f"Test files created in: {os.path.dirname(adjusted_file)}")
        print(f"Output file: {output_file}")
        
        # Show success message
        root = tk.Tk()
        root.withdraw()  # Hide the root window
        messagebox.showinfo("Tests Completed", f"All tests completed successfully!\nOutput file: {output_file}")
        root.destroy()
        
    except Exception as e:
        print(f"Error during testing: {str(e)}")
        
        # Show error message
        root = tk.Tk()
        root.withdraw()  # Hide the root window
        messagebox.showerror("Test Error", f"Error during testing: {str(e)}")
        root.destroy()

if __name__ == "__main__":
    main() 