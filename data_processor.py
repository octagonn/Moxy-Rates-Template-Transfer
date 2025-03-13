#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Data Processor module for Moxy Rates Template Transfer

This module provides functionality for loading, transforming, and integrating
Excel data between Adjusted Rates and Template files.
"""

import os
import logging
import pandas as pd


class DataProcessor:
    """Handles Excel data processing operations."""
    
    def __init__(self):
        """Initialize the data processor."""
        logging.info("DataProcessor initialized")
    
    def load_excel_file(self, file_path, sheet_name=None):
        """
        Load an Excel file into a pandas DataFrame.
        
        Args:
            file_path: Path to the Excel file
            sheet_name: Name of the sheet to load (optional)
            
        Returns:
            DataFrame: Loaded data
        """
        try:
            if sheet_name:
                df = pd.read_excel(file_path, sheet_name=sheet_name)
                logging.info(f"Loaded Excel file '{file_path}', sheet '{sheet_name}' with {len(df)} rows and {len(df.columns)} columns")
            else:
                df = pd.read_excel(file_path)
                logging.info(f"Loaded Excel file '{file_path}' with {len(df)} rows and {len(df.columns)} columns")
                
            return df
        except Exception as e:
            error_msg = f"Error loading Excel file '{file_path}': {str(e)}"
            logging.error(error_msg)
            raise ValueError(error_msg)
    
    def transform_data(self, adjusted_rates_df):
        """
        Transform the adjusted rates data to match template format.
        
        Args:
            adjusted_rates_df: DataFrame containing the standardized adjusted rates data
            
        Returns:
            DataFrame: Transformed data ready for template
        """
        try:
            logging.info("Transforming data...")
            
            # Check that required columns are present
            required_columns = [
                'Coverage', 'Term', 'Miles', 'FromMiles', 'ToMiles', 
                'MinYear', 'MaxYears', 'Class', 'RateCost', 'Deductible'
            ]
            
            missing = [col for col in required_columns if col not in adjusted_rates_df.columns]
            if missing:
                error_msg = f"Missing required columns for transformation: {', '.join(missing)}"
                logging.error(error_msg)
                raise ValueError(error_msg)
            
            # Make a copy of the data to avoid modifying the original
            df = adjusted_rates_df.copy()
            
            # Pivot the data to create separate columns for each deductible value
            logging.info("Pivoting data by Deductible")
            adjusted_rates_pivot = df.pivot_table(
                index=['Coverage', 'Term', 'Miles', 'FromMiles', 'ToMiles', 'MinYear', 'MaxYears', 'Class'],
                columns='Deductible',
                values='RateCost',
                aggfunc='first'
            ).reset_index()
            
            # Clean up the pivot columns
            adjusted_rates_pivot.columns.name = None
            
            # Create columns for each deductible value
            numeric_columns = [col for col in adjusted_rates_pivot.columns if isinstance(col, (int, float))]
            
            # Create column mapping for deductible values
            column_mapping = {}
            for col in numeric_columns:
                column_mapping[col] = f"Deduct{int(col)}"
            
            # Rename columns
            adjusted_rates_pivot = adjusted_rates_pivot.rename(columns=column_mapping)
            
            # Add PlanDeduct column
            deductible_columns = [col for col in adjusted_rates_pivot.columns if 'Deduct' in str(col)]
            
            logging.info(f"Creating PlanDeduct column from {len(deductible_columns)} deductible columns")
            
            # Fill PlanDeduct column with the minimum available deductible for each row
            adjusted_rates_pivot['PlanDeduct'] = adjusted_rates_pivot[deductible_columns].apply(
                lambda x: self._get_min_deductible(x, deductible_columns), axis=1
            )
            
            # Round all numeric columns to 2 decimal places
            for col in deductible_columns:
                adjusted_rates_pivot[col] = adjusted_rates_pivot[col].round(2)
            
            # Clean up any None values
            adjusted_rates_pivot = adjusted_rates_pivot.fillna('')
            
            logging.info(f"Transformation complete. Result has {len(adjusted_rates_pivot)} rows and {len(adjusted_rates_pivot.columns)} columns")
            return adjusted_rates_pivot
            
        except Exception as e:
            error_msg = f"Error transforming data: {str(e)}"
            logging.error(error_msg, exc_info=True)
            raise ValueError(error_msg)
    
    def _get_min_deductible(self, row, deductible_columns):
        """
        Get the minimum available deductible from a row.
        
        Args:
            row: DataFrame row
            deductible_columns: List of deductible column names
            
        Returns:
            int: Minimum deductible value
        """
        # Priority for deductible 100 if it exists and has a value
        if 'Deduct100' in deductible_columns and pd.notna(row.get('Deduct100')):
            return 100
        
        # Otherwise, find the minimum available deductible
        available_deducts = []
        for col in deductible_columns:
            if pd.notna(row.get(col)) and row.get(col) != '':
                try:
                    deduct_value = int(col.replace('Deduct', ''))
                    available_deducts.append(deduct_value)
                except (ValueError, AttributeError):
                    pass
        
        return min(available_deducts) if available_deducts else 0
    
    def integrate_with_template(self, transformed_data, template_df):
        """
        Integrate the transformed data with the template.
        
        Args:
            transformed_data: DataFrame containing the transformed data
            template_df: DataFrame containing the template
            
        Returns:
            DataFrame: Updated template with integrated data
        """
        try:
            logging.info("Integrating data with template...")
            
            # Check if template has any rows
            if len(template_df) == 0:
                logging.info("Empty template, using transformed data directly")
                return transformed_data
            
            # Check if template has all required columns
            missing_columns = [col for col in transformed_data.columns 
                              if col not in template_df.columns]
            
            # If template is missing columns, add them
            template_copy = template_df.copy()
            if missing_columns:
                logging.info(f"Adding missing columns to template: {', '.join(missing_columns)}")
                for col in missing_columns:
                    template_copy[col] = None
            
            # If template has some placeholder data, we might want to remove it
            # Check if this is a template with placeholder rows (typically < 5 rows)
            if len(template_copy) < 5:
                placeholder_rows = len(template_copy)
                
                # Combine template (for structure) with transformed data (for content)
                logging.info(f"Template appears to be a structure template with {placeholder_rows} placeholder rows")
                final_df = pd.concat([template_copy.iloc[0:0], transformed_data], ignore_index=True)
                
                # Copy column order from template
                column_order = [col for col in template_copy.columns if col in final_df.columns]
                extra_cols = [col for col in final_df.columns if col not in column_order]
                final_df = final_df[column_order + extra_cols]
                
                logging.info(f"Integrated data has {len(final_df)} rows and {len(final_df.columns)} columns")
                return final_df
            else:
                # Template seems to already have data, so append the new data
                logging.info(f"Template already has {len(template_copy)} rows of data, appending transformed data")
                final_df = pd.concat([template_copy, transformed_data], ignore_index=True)
                
                # Remove duplicate rows, keeping the transformed data version if duplicates exist
                # Identify key columns for duplication checking
                key_cols = [col for col in ['Coverage', 'Term', 'Miles', 'FromMiles', 'ToMiles', 'Class', 'Deductible']
                           if col in final_df.columns]
                
                if key_cols:
                    before_dedup = len(final_df)
                    final_df = final_df.drop_duplicates(subset=key_cols, keep='last')
                    after_dedup = len(final_df)
                    
                    if before_dedup > after_dedup:
                        logging.info(f"Removed {before_dedup - after_dedup} duplicate rows")
                
                logging.info(f"Final integrated data has {len(final_df)} rows and {len(final_df.columns)} columns")
                return final_df
                
        except Exception as e:
            error_msg = f"Error integrating with template: {str(e)}"
            logging.error(error_msg, exc_info=True)
            raise ValueError(error_msg) 