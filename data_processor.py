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
import openpyxl


class DataProcessor:
    """Handles Excel data processing operations."""
    
    def __init__(self):
        """Initialize the data processor."""
        logging.info("DataProcessor initialized")
        self.default_deductible = "100"  # Default value, can be changed by user
    
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
            logging.info(f"Loading Excel file: {file_path}")
            
            # Special handling for template files to preserve unnamed columns
            is_template = 'template' in str(file_path).lower()
            
            if is_template:
                logging.info("Detected template file - using special handling to preserve all columns")
                # For template files:
                # 1. keep_default_na=False prevents pandas from converting empty cells to NaN
                # 2. na_values=[] prevents any values from being interpreted as NaN
                # 3. dtype=str forces all columns to be string type, preventing numeric conversion
                df = pd.read_excel(
                    file_path,
                    sheet_name=sheet_name,
                    keep_default_na=False,
                    na_values=[],
                    dtype=str
                )
                
                # Get original column names exactly as they appear in Excel
                xl = pd.ExcelFile(file_path)
                sheet = xl.book.active if sheet_name is None else xl.book[sheet_name]
                
                # Extract the header row (usually row 1)
                header_values = [cell.value for cell in sheet[1]]
                
                # Update column names to match Excel exactly
                if len(header_values) == len(df.columns):
                    # Replace None with empty string but preserve column headers
                    header_values = ['' if v is None else v for v in header_values]
                    df.columns = header_values
                    logging.info(f"Preserved exact column headers from template: {header_values}")
                else:
                    logging.warning(f"Column count mismatch between Excel headers ({len(header_values)}) and DataFrame ({len(df.columns)})")
            else:
                # For non-template files, use standard pandas behavior but still preserve empty cells
                df = pd.read_excel(
                    file_path,
                    sheet_name=sheet_name,
                    keep_default_na=False,
                    na_values=[]
                )
            
            logging.info(f"Loaded {len(df)} rows and {len(df.columns)} columns from {file_path}")
            logging.info(f"Column names: {df.columns.tolist()}")
            
            return df
            
        except Exception as e:
            logging.error(f"Error loading Excel file {file_path}: {str(e)}", exc_info=True)
            # Return empty DataFrame
            return pd.DataFrame()
    
    def transform_data(self, source_df, mapping):
        """
        Transform the data from the source format to the template format.
        
        Args:
            source_df (DataFrame): The source data
            mapping (dict): The mapping from source to template columns
            
        Returns:
            DataFrame: Transformed data
        """
        try:
            logging.info("STEP 1: Starting data transformation process")
            logging.info(f"Source data shape: {source_df.shape}")
            logging.info(f"Source columns: {source_df.columns.tolist()}")
            logging.info(f"Mapping: {mapping}")
            
            # Safety check - return source df if empty
            if source_df.empty:
                logging.warning("Source dataframe is empty, returning as is")
                return source_df
            
            # Log a sample row for debugging
            if len(source_df) > 0:
                logging.info(f"Sample source row: {source_df.iloc[0].to_dict()}")
            
            # STEP 2: Check for Deductible and RateCost in mapping
            if "Deductible" not in mapping or "RateCost" not in mapping:
                logging.error("Missing required mapping for pivot operations: Deductible and/or RateCost")
                deductible_cols = [col for col in mapping.keys() if "deduct" in str(col).lower()]
                cost_cols = [col for col in mapping.keys() if any(x in str(col).lower() for x in ["rate", "cost", "premium"])]
                
                logging.info(f"Potential deductible columns: {deductible_cols}")
                logging.info(f"Potential rate/cost columns: {cost_cols}")
                
                # Try to auto-detect if possible
                if "Deductible" not in mapping and deductible_cols:
                    mapping["Deductible"] = deductible_cols[0]
                    logging.info(f"Auto-assigned Deductible mapping to {deductible_cols[0]}")
                
                if "RateCost" not in mapping and cost_cols:
                    mapping["RateCost"] = cost_cols[0]
                    logging.info(f"Auto-assigned RateCost mapping to {cost_cols[0]}")
                
                # Check again after auto-detection
                if "Deductible" not in mapping or "RateCost" not in mapping:
                    logging.warning("Cannot perform pivot operation due to missing mappings. Returning source data.")
                    return source_df
            
            # STEP 3: Create inverse mapping from template fields to source columns
            logging.info("STEP 3: Creating mapping from template fields to source columns")
            inverse_mapping = {}
            for field, source_col in mapping.items():
                if source_col and field and source_col in source_df.columns:
                    inverse_mapping[field] = source_col
            
            # Check if we have the essential mappings
            source_deductible_col = inverse_mapping.get("Deductible")
            source_rate_cost_col = inverse_mapping.get("RateCost")
            
            if not source_deductible_col or not source_rate_cost_col:
                logging.error(f"Missing essential columns: Deductible={source_deductible_col}, RateCost={source_rate_cost_col}")
                return source_df
            
            # STEP 4: Create a new DataFrame with renamed columns
            logging.info("STEP 4: Renaming columns according to mapping")
            renamed_df = pd.DataFrame()
            
            # Copy columns with their new names
            for template_field, source_col in inverse_mapping.items():
                if source_col in source_df.columns:
                    renamed_df[template_field] = source_df[source_col].fillna('')
                    logging.info(f"Renamed column {source_col} to {template_field}")
            
            # STEP 5: Prepare for pivoting
            logging.info("STEP 5: Preparing for pivot operation")
            
            # Determine columns to group by
            group_cols = [col for col in renamed_df.columns 
                         if col not in ['Deductible', 'RateCost', 'PlanDeduct']]
            
            # Check if we have necessary columns for pivoting
            if 'Deductible' not in renamed_df.columns or 'RateCost' not in renamed_df.columns:
                logging.error("Missing required columns for pivot. Returning renamed dataframe with original columns.")
                return renamed_df
            
            logging.info(f"Will group by these columns for pivoting: {group_cols}")
            
            # Log sample data before pivoting
            if len(renamed_df) > 0:
                logging.info(f"Sample data before pivoting: {renamed_df.iloc[0].to_dict()}")
            
            # STEP 6: Prepare dictionaries for the pivot
            logging.info("STEP 6: Executing pivot operation")
            
            try:
                # Create a dictionary to store the grouped data
                grouped_data = {}
                
                # Process each row
                for idx, row in renamed_df.iterrows():
                    # Create a key from the group columns
                    key_parts = []
                    for col in group_cols:
                        val = row.get(col)
                        # For Series objects (can happen with duplicate column names)
                        if isinstance(val, pd.Series):
                            val = val.iloc[0] if not val.empty else None
                        key_parts.append(str(val) if pd.notna(val) else "")
                    
                    key = "||".join(key_parts)
                    
                    # Get deductible and rate cost
                    try:
                        deductible = str(row['Deductible']).strip() if pd.notna(row.get('Deductible')) else ""
                        rate_cost = row.get('RateCost') if pd.notna(row.get('RateCost')) else None
                        
                        # Handle Series objects
                        if isinstance(deductible, pd.Series):
                            deductible = str(deductible.iloc[0]).strip() if not deductible.empty else ""
                        if isinstance(rate_cost, pd.Series):
                            rate_cost = rate_cost.iloc[0] if not rate_cost.empty else None
                    except Exception as e:
                        logging.warning(f"Error getting deductible or rate cost: {str(e)}")
                        continue
                    
                    # Skip rows with empty values
                    if not deductible or deductible.lower() == 'nan' or rate_cost is None:
                        logging.info(f"Skipping row with invalid deductible: {deductible} or missing rate cost")
                        continue
                    
                    # If this key doesn't exist, create a new entry
                    if key not in grouped_data:
                        entry = {}
                        # Add the group column values
                        for i, col in enumerate(group_cols):
                            entry[col] = key_parts[i]
                        
                        grouped_data[key] = entry
                        logging.info(f"Created new entry for key: {key}")
                    
                    # Add the deductible value column 
                    # Clean up the deductible value
                    deductible_clean = ''.join(c for c in deductible if c.isdigit())
                    if not deductible_clean:
                        logging.warning(f"Deductible '{deductible}' has no numeric characters, skipping")
                        continue
                    
                    # Create the deductible column name following exact format: "Deduct50", "Deduct100", etc.
                    deduct_col = f"Deduct{deductible_clean}"
                    logging.info(f"Created deductible column: '{deduct_col}' from value '{deductible}'")
                    
                    # Add the rate cost to the appropriate deductible column
                    grouped_data[key][deduct_col] = rate_cost
                    logging.info(f"Added {deduct_col}={rate_cost} to key: {key}")
                
                # Convert the dictionary to a DataFrame
                result_df = pd.DataFrame(list(grouped_data.values()))
                
                if result_df.empty:
                    logging.warning("Pivot produced no data, returning renamed DataFrame")
                    return renamed_df
                
                logging.info(f"Created pivoted dataframe with shape: {result_df.shape}")
                logging.info(f"Pivoted columns: {result_df.columns.tolist()}")
                
            except Exception as e:
                logging.error(f"Error during pivoting: {str(e)}", exc_info=True)
                logging.warning("Using alternative pivot method due to error")
                # Fallback method if the above fails
                result_df = renamed_df
            
            # STEP 7: Initialize any missing required columns with empty string
            required_cols = [
                'CompanyCode', 'Term', 'Miles', 'FromMiles', 'ToMiles', 'Coverage', 
                'State', 'Class', 'PlanDeduct', 'Markup', 'New/Used', 'MaxYears', 
                'SurchargeCode', 'PlanCode', 'RateCardCode', 'ClassListCode', 
                'MinYear', 'IncScCode', 'IncScAmt'
            ]
            
            for col in required_cols:
                if col not in result_df.columns:
                    result_df[col] = ''
            
            # Initialize standard deductible columns with empty string if missing
            standard_deducts = ['Deduct0', 'Deduct50', 'Deduct100', 'Deduct200', 'Deduct250', 'Deduct500']
            for deduct_col in standard_deducts:
                if deduct_col not in result_df.columns:
                    result_df[deduct_col] = ''
            
            # Ensure all empty values are properly set to empty string
            for col in result_df.columns:
                result_df[col] = result_df[col].fillna('')
            
            logging.info(f"Final transformed data shape: {result_df.shape}")
            logging.info(f"Final columns: {result_df.columns.tolist()}")
            
            return result_df
            
        except Exception as e:
            logging.error(f"Error in data transformation: {str(e)}", exc_info=True)
            return source_df
    
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
        
        # Fix: Ensure we're properly checking if the list is empty and using min() correctly
        if not available_deducts:
            return 0
        else:
            return min(available_deducts)  # This calls the min function properly
    
    def integrate_with_template(self, transformed_data, template_path):
        """
        Integrate the transformed data with the template.
        
        Args:
            transformed_data (DataFrame): The transformed data
            template_path (str): Path to the template file
            
        Returns:
            DataFrame: Data integrated with template format
        """
        try:
            logging.info("Starting template integration process")
            
            # Initialize template columns with empty strings
            template_columns = [
                'CompanyCode', 'Term', 'Miles', 'FromMiles', 'ToMiles', 'Coverage', 
                'State', 'Class', 'PlanDeduct', 'Deduct0', 'Deduct50', 'Deduct100', 
                'Deduct200', 'Deduct250', 'Deduct500', 'Markup', 'New/Used', 'MaxYears', 
                'SurchargeCode', 'PlanCode', 'RateCardCode', 'ClassListCode', 'MinYear', 
                'IncScCode', 'IncScAmt'
            ]
            
            # Create a new DataFrame with template columns
            result_df = pd.DataFrame(columns=template_columns)
            
            # Copy data from transformed_data, filling missing values with empty string
            for col in template_columns:
                if col in transformed_data.columns:
                    result_df[col] = transformed_data[col].fillna('')
                else:
                    result_df[col] = ''
            
            # Ensure all columns are string type
            for col in result_df.columns:
                result_df[col] = result_df[col].astype(str)
                # Replace 'nan' and 'None' strings with empty string
                result_df[col] = result_df[col].replace({'nan': '', 'None': '', 'NaN': ''})
            
            logging.info(f"Final integrated data shape: {result_df.shape}")
            logging.info(f"Final columns: {result_df.columns.tolist()}")
            
            return result_df
            
        except Exception as e:
            logging.error(f"Error in template integration: {str(e)}", exc_info=True)
            return transformed_data

    def _add_plan_deduct_column(self, df):
        """
        Add the PlanDeduct column to the DataFrame and ensure proper column ordering.
        
        Args:
            df (DataFrame): The DataFrame with deductible columns
            
        Returns:
            DataFrame: DataFrame with PlanDeduct column added and columns properly ordered
        """
        try:
            if df.empty:
                return df
            
            # Find all deductible columns
            deduct_columns = [col for col in df.columns if str(col).startswith('Deduct') and col != 'PlanDeduct']
            
            if not deduct_columns:
                logging.warning("No deductible columns found, cannot add PlanDeduct")
                return df
            
            # Extract deductible values from column names
            deduct_values = []
            for col in deduct_columns:
                try:
                    val = ''.join(filter(str.isdigit, str(col)))
                    if val:
                        deduct_values.append((int(val), col))
                except ValueError:
                    pass
            
            # Sort by deductible value
            deduct_values.sort()
            logging.info(f"Sorted deductible values: {deduct_values}")
            
            # Add PlanDeduct column if it doesn't exist
            if 'PlanDeduct' not in df.columns:
                df['PlanDeduct'] = None
            
            # Process each row
            for idx in range(len(df)):
                # Following the example in the image, we'll set specific values for PlanDeduct
                
                # 1. If default_deductible is set, check if that column has a value
                if hasattr(self, 'default_deductible') and self.default_deductible:
                    default_col = f"Deduct{self.default_deductible}"
                    if default_col in df.columns and pd.notna(df.loc[idx, default_col]):
                        df.loc[idx, 'PlanDeduct'] = self.default_deductible
                        continue
                
                # 2. If Class is C or D, use lowest deductible (usually 0)
                if 'Class' in df.columns and pd.notna(df.loc[idx, 'Class']):
                    class_val = str(df.loc[idx, 'Class']).strip().upper()
                    if class_val in ['C', 'D']:
                        # Find the lowest deductible column with a value
                        for val, col in deduct_values:
                            if pd.notna(df.loc[idx, col]):
                                df.loc[idx, 'PlanDeduct'] = str(val)
                                break
                        continue
                
                # 3. For all other classes (E, F, G, H), use deductible 100 if available
                default_val = "100"
                default_col = f"Deduct{default_val}"
                if default_col in df.columns and pd.notna(df.loc[idx, default_col]):
                    df.loc[idx, 'PlanDeduct'] = default_val
                    continue
                
                # 4. Fallback to lowest available deductible
                for val, col in deduct_values:
                    if pd.notna(df.loc[idx, col]):
                        df.loc[idx, 'PlanDeduct'] = str(val)
                        break
            
            # Organize columns in the desired order based on the second image example
            # Define the expected column order following the second image example
            # This is a template of the order we want (not all columns may exist)
            desired_order = [
                'CompanyCode', 'Term', 'Miles', 'FromMiles', 'ToMiles', 'Coverage', 'State', 'Class', 
                'PlanDeduct'
            ]
            
            # Add all deductible columns in numerical order
            sorted_deduct_cols = [col for _, col in deduct_values]
            
            # Add remaining columns in the desired order
            remaining_order = [
                'Markup', 'New/Used', 'MaxYears', 'SurchargeCode', 'PlanCode', 
                'RateCardCode', 'ClassListCode', 'MinYear', 'IncSCode', 'IncSAmt'
            ]
            
            # Create the final order based on what's actually in the DataFrame
            final_order = []
            
            # First add columns from desired_order that exist in the DataFrame
            for col in desired_order:
                if col in df.columns:
                    final_order.append(col)
            
            # Next add the deductible columns in sorted order
            for col in sorted_deduct_cols:
                if col in df.columns and col not in final_order:
                    final_order.append(col)
            
            # Next add remaining columns from the predefined order
            for col in remaining_order:
                if col in df.columns and col not in final_order:
                    final_order.append(col)
            
            # Finally add any other columns not yet included
            for col in df.columns:
                if col not in final_order:
                    final_order.append(col)
            
            logging.info(f"Reordering columns to match desired format: {final_order}")
            
            # Apply the new column order
            if final_order:
                df = df[final_order]
            
            return df
            
        except Exception as e:
            logging.error(f"Error adding PlanDeduct column: {str(e)}", exc_info=True)
            return df 

    def save_excel_file(self, df, output_file, sheet_name="Sheet1"):
        """
        Save DataFrame to Excel with proper handling of empty columns.
        
        Args:
            df (DataFrame): The DataFrame to save
            output_file (str): Path to save the Excel file
            sheet_name (str): Name of the sheet to save to
        """
        try:
            logging.info(f"Saving DataFrame to {output_file}")
            logging.info(f"DataFrame shape: {df.shape}")
            logging.info(f"Columns: {df.columns.tolist()}")
            
            # Create Excel writer with openpyxl engine
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                # Replace empty strings with None for proper Excel empty cells
                df_to_save = df.replace('', None)
                
                # Write DataFrame without index
                df_to_save.to_excel(writer, sheet_name=sheet_name, index=False)
                
                # Auto-adjust column widths
                worksheet = writer.sheets[sheet_name]
                for idx, col in enumerate(df_to_save.columns):
                    # Get maximum length of column name and its contents
                    max_length = max(
                        df_to_save[col].astype(str).apply(len).max(),
                        len(str(col))
                    )
                    # Add a little extra space
                    adjusted_width = (max_length + 2)
                    # Set column width
                    worksheet.column_dimensions[chr(65 + idx)].width = adjusted_width
            
            logging.info(f"Successfully saved DataFrame to {output_file}")
            
        except Exception as e:
            logging.error(f"Error saving DataFrame to Excel: {str(e)}", exc_info=True)
            raise 