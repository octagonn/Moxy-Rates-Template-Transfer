#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
File Analyzer module for Moxy Rates Template Transfer

This module provides functionality to analyze Excel files and detect their structure,
including column names, data types, and potential mapping suggestions.
"""

import os
import logging
import pandas as pd
from fuzzywuzzy import fuzz

class FileAnalyzer:
    """Analyzes Excel files to detect structure and suggest mappings."""
    
    def __init__(self):
        """Initialize the file analyzer."""
        logging.info("FileAnalyzer initialized")
    
    def get_sheet_names(self, file_path):
        """
        Get a list of sheet names from an Excel file.
        
        Args:
            file_path: Path to the Excel file
            
        Returns:
            list: List of sheet names
        """
        try:
            xls = pd.ExcelFile(file_path)
            return xls.sheet_names
        except Exception as e:
            logging.error(f"Error reading sheet names: {str(e)}")
            raise ValueError(f"Unable to read Excel file: {str(e)}")
    
    def analyze_file_structure(self, excel_path, sheet_name=None):
        """
        Analyze the structure of an Excel file to identify its format.
        
        Args:
            excel_path: Path to the Excel file
            sheet_name: Optional specific sheet to analyze
            
        Returns:
            dict: Structure information including column details, data types,
                  and potential mapping suggestions
        """
        try:
            logging.info(f"Analyzing file structure: {excel_path}, sheet: {sheet_name}")
            
            # Load the Excel file
            if sheet_name:
                df = pd.read_excel(excel_path, sheet_name=sheet_name)
                structure = self.analyze_sheet_structure(df)
                structure['file_path'] = excel_path
                structure['sheet_name'] = sheet_name
                return structure
            else:
                # Try to analyze all sheets if none specified
                xls = pd.ExcelFile(excel_path)
                sheets_analysis = {}
                
                for sheet in xls.sheet_names:
                    sheets_analysis[sheet] = self.analyze_sheet_structure(
                        pd.read_excel(excel_path, sheet_name=sheet))
                
                # Find the most likely main data sheet
                main_sheet = self.identify_main_data_sheet(sheets_analysis)
                
                return {
                    'file_path': excel_path,
                    'sheets': sheets_analysis,
                    'suggested_main_sheet': main_sheet,
                    'sheet_name': main_sheet,  # For compatibility
                    'columns': sheets_analysis[main_sheet]['columns'] if main_sheet else {},
                    'column_mapping_suggestions': sheets_analysis[main_sheet]['column_mapping_suggestions'] if main_sheet else {}
                }
            
        except Exception as e:
            logging.error(f"Error analyzing file structure: {str(e)}", exc_info=True)
            return {'error': str(e)}
    
    def analyze_sheet_structure(self, df):
        """
        Analyze the structure of a dataframe to identify columns and data types.
        
        Args:
            df: Pandas DataFrame to analyze
            
        Returns:
            dict: Structure information including column details and data types
        """
        structure = {
            'columns': {},
            'row_count': len(df),
            'potential_key_columns': [],
            'data_summary': {}
        }
        
        logging.info(f"Analyzing sheet with {len(df.columns)} columns and {len(df)} rows")
        
        # Analyze each column
        for col in df.columns:
            try:
                # Get sample values, handling empty columns
                sample_values = []
                if not df[col].empty:
                    sample_values = df[col].dropna().sample(
                        min(5, len(df[col].dropna()))).tolist()
                
                col_info = {
                    'data_type': str(df[col].dtype),
                    'unique_values': df[col].nunique(),
                    'sample_values': sample_values,
                    'null_percentage': (df[col].isna().sum() / len(df)) * 100 if len(df) > 0 else 0
                }
                
                # Detect if column is likely a key/identifier
                if col_info['unique_values'] > (len(df) * 0.9) and len(df) > 10:
                    structure['potential_key_columns'].append(col)
                    
                # Detect if column has numeric values
                if df[col].dtype in ['int64', 'float64']:
                    if not df[col].empty:
                        col_info['min'] = df[col].min()
                        col_info['max'] = df[col].max()
                        col_info['mean'] = df[col].mean()
                    
                    # Detect if column might be a deductible column
                    if 'deduct' in str(col).lower() or (
                        not df[col].empty and 
                        df[col].min() >= 0 and 
                        df[col].max() <= 1000):
                        col_info['possible_type'] = 'deductible'
                        
                # Detect if column might be a year column
                if df[col].dtype in ['int64'] and not df[col].empty:
                    if df[col].min() >= 1990 and df[col].max() <= 2050:
                        col_info['possible_type'] = 'year'
                        
                # Look for common column name patterns
                col_lower = str(col).lower()
                if any(term in col_lower for term in ['mile', 'distance', 'km']):
                    col_info['possible_type'] = 'mileage'
                elif any(term in col_lower for term in ['from', 'start', 'min']):
                    if 'mile' in col_lower:
                        col_info['possible_type'] = 'frommiles'
                elif any(term in col_lower for term in ['to', 'end', 'max']):
                    if 'mile' in col_lower:
                        col_info['possible_type'] = 'tomiles'
                elif any(term in col_lower for term in ['cover', 'coverage']):
                    col_info['possible_type'] = 'coverage'
                elif any(term in col_lower for term in ['class', 'category', 'type']):
                    col_info['possible_type'] = 'class'
                elif any(term in col_lower for term in ['term', 'duration', 'period']):
                    col_info['possible_type'] = 'term'
                elif any(term in col_lower for term in ['rate', 'cost', 'price']):
                    col_info['possible_type'] = 'ratecost'
                elif 'minyear' in col_lower or ('min' in col_lower and 'year' in col_lower):
                    col_info['possible_type'] = 'minyear'
                elif 'maxyear' in col_lower or ('max' in col_lower and 'year' in col_lower):
                    col_info['possible_type'] = 'maxyears'
                    
                structure['columns'][col] = col_info
                
            except Exception as e:
                logging.error(f"Error analyzing column {col}: {str(e)}")
                structure['columns'][col] = {
                    'error': str(e),
                    'data_type': str(df[col].dtype) if col in df else 'unknown'
                }
        
        # Detect if this is likely a template or data sheet
        structure['likely_purpose'] = 'template' if len(df) < 5 else 'data'
        
        # Generate column similarity scores for common required fields
        structure['column_mapping_suggestions'] = self.suggest_column_mappings(structure['columns'])
        
        return structure
    
    def identify_main_data_sheet(self, sheets_analysis):
        """
        Identify the most likely sheet containing main data.
        
        Args:
            sheets_analysis: Dictionary of sheet analysis results
            
        Returns:
            str: Name of the most likely main data sheet
        """
        # Criteria for selecting the main data sheet
        best_score = 0
        best_sheet = None
        
        for sheet_name, analysis in sheets_analysis.items():
            score = 0
            
            # More rows is better
            score += min(100, analysis.get('row_count', 0) / 10)
            
            # More columns that match our required fields is better
            score += len(analysis.get('column_mapping_suggestions', {})) * 10
            
            # Keywords in sheet name
            if any(term in sheet_name.lower() for term in ['rate', 'data', 'cost', 'deductible', 'miles']):
                score += 50
                
            # Specific sheet names
            if sheet_name.lower() in ['dealer cost rates', 'rates', 'data']:
                score += 100
                
            # If it has more than 10 columns, likely a data sheet
            if len(analysis.get('columns', {})) > 10:
                score += 50
                
            # If it has Coverage, Miles, and Term columns, likely our target
            required_columns = ['coverage', 'miles', 'term', 'deductible', 'class']
            suggestions = analysis.get('column_mapping_suggestions', {})
            found_required = [field for field in required_columns if field in suggestions]
            score += len(found_required) * 20
            
            logging.debug(f"Sheet '{sheet_name}' score: {score}")
            
            if score > best_score:
                best_score = score
                best_sheet = sheet_name
                
        logging.info(f"Identified main sheet: {best_sheet} with score {best_score}")
        return best_sheet
    
    def suggest_column_mappings(self, columns_info):
        """
        Suggest possible mappings for columns based on names and data characteristics.
        
        Args:
            columns_info: Dictionary of column information
            
        Returns:
            dict: Suggested mappings for required fields
        """
        required_fields = {
            'coverage': ['coverage', 'cover', 'cov', 'protection', 'plan'],
            'term': ['term', 'duration', 'period', 'months', 'length'],
            'miles': ['miles', 'mileage', 'distance', 'odometer', 'mi'],
            'frommiles': ['frommiles', 'from_miles', 'startmiles', 'minmiles', 'lowmiles', 'from mile'],
            'tomiles': ['tomiles', 'to_miles', 'endmiles', 'maxmiles', 'highmiles', 'to mile'],
            'minyear': ['minyear', 'min_year', 'yearmin', 'startyear', 'fromyear', 'vehicleyearmin'],
            'maxyears': ['maxyears', 'max_years', 'yearmax', 'endyear', 'toyear', 'vehicleyearmax', 'year limit'],
            'class': ['class', 'vehicleclass', 'category', 'tier', 'classification'],
            'ratecost': ['ratecost', 'rate', 'cost', 'price', 'premium', 'dealer', 'dealercost'],
            'deductible': ['deductible', 'deduct', 'ded', 'deductable', 'deduction']
        }
        
        logging.info(f"Generating column mapping suggestions for {len(columns_info)} columns")
        suggestions = {}
        
        # Use fuzzy matching to find the best match for each required field
        for required_field, synonyms in required_fields.items():
            best_match = None
            best_score = 0
            match_reason = ""
            
            for col_name in columns_info.keys():
                col_lower = str(col_name).lower()
                col_info = columns_info[col_name]
                
                # Check exact matches first
                if col_lower == required_field:
                    best_match = col_name
                    best_score = 100
                    match_reason = "Exact name match"
                    break
                    
                # Check synonyms
                for synonym in synonyms:
                    if synonym in col_lower:
                        score = 80  # Good match based on substring
                        if score > best_score:
                            best_match = col_name
                            best_score = score
                            match_reason = f"Contains synonym '{synonym}'"
                
                # Fuzzy matching if no direct match
                if best_score < 80:
                    for synonym in synonyms + [required_field]:
                        ratio = fuzz.ratio(col_lower, synonym)
                        partial_ratio = fuzz.partial_ratio(col_lower, synonym)
                        token_sort_ratio = fuzz.token_sort_ratio(col_lower, synonym)
                        
                        # Use the best fuzzy match score
                        score = max(ratio, partial_ratio, token_sort_ratio)
                        
                        if score > best_score and score > 60:  # Threshold for fuzzy matching
                            best_match = col_name
                            best_score = score
                            match_reason = f"Fuzzy match with '{synonym}' (score: {score})"
                            
                # Check for possible_type field match
                if 'possible_type' in col_info:
                    if col_info['possible_type'] == required_field:
                        score = 90  # High confidence based on detected data type
                        if score > best_score:
                            best_match = col_name
                            best_score = score
                            match_reason = "Data type detection"
            
            if best_match:
                suggestions[required_field] = {
                    'suggested_column': best_match,
                    'confidence': best_score,
                    'reason': match_reason
                }
                logging.debug(f"Mapping suggestion: {required_field} -> {best_match} (confidence: {best_score})")
                
        return suggestions
    
    def identify_column_purpose(self, column_name, sample_values):
        """
        Use heuristic approaches to identify the purpose of a column.
        
        Args:
            column_name: Name of the column
            sample_values: Sample values from the column
            
        Returns:
            tuple: (purpose, confidence_score)
        """
        # Normalize column name
        name = str(column_name).lower().replace(' ', '').replace('_', '')
        
        # Check for exact matches in common column names
        column_patterns = {
            'coverage': ['coverage', 'coveragetype', 'cover', 'protection', 'plan'],
            'term': ['term', 'termmonths', 'termlength', 'months', 'duration'],
            'miles': ['miles', 'mileage', 'odometer', 'distance', 'milelimit'],
            'frommiles': ['frommiles', 'startmiles', 'minmiles', 'lowmiles'],
            'tomiles': ['tomiles', 'endmiles', 'maxmiles', 'highmiles'],
            'minyear': ['minyear', 'yearmin', 'startyear', 'fromyear', 'vehicleyearmin'],
            'maxyears': ['maxyears', 'yearmax', 'maxyear', 'toyear', 'vehicleyearmax'],
            'class': ['class', 'vehicleclass', 'category', 'tier', 'classification'],
            'ratecost': ['ratecost', 'rate', 'cost', 'price', 'premium', 'dealer', 'dealercost'],
            'deductible': ['deductible', 'deduct', 'ded', 'deductibleamount']
        }
        
        # Check for pattern matches
        for purpose, patterns in column_patterns.items():
            if any(pattern == name for pattern in patterns):
                return purpose, 100  # Exact match
            if any(pattern in name for pattern in patterns):
                return purpose, 80   # Partial match
        
        # Analyze sample values if no match by name
        if sample_values:
            # Convert to strings for analysis
            str_values = [str(val) for val in sample_values if val is not None]
            
            if str_values:
                # Check for deductible patterns
                if all(str(val).isdigit() and int(val) in [0, 50, 100, 200, 250, 500, 1000] 
                      for val in str_values if val is not None and str(val).isdigit()):
                    return 'deductible', 70
                    
                # Check for mileage patterns (typically multiples of 1000)
                if all(str(val).isdigit() and int(val) % 1000 == 0 and int(val) > 1000 
                      for val in str_values if val is not None and str(val).isdigit()):
                    return 'miles', 60
                    
                # Check for term patterns (typically 12, 24, 36, 48, 60, 72, 84)
                if all(str(val).isdigit() and int(val) in [12, 24, 36, 48, 60, 72, 84, 96, 120] 
                      for val in str_values if val is not None and str(val).isdigit()):
                    return 'term', 70
                    
                # Check for year patterns
                if all(str(val).isdigit() and int(val) >= 1990 and int(val) <= 2030 
                      for val in str_values if val is not None and str(val).isdigit()):
                    return 'minyear', 60  # Could be either min or max year
        
        # No clear match
        return 'unknown', 0 