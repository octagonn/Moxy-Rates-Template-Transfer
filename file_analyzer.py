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
import re

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
    
    def analyze_file_structure(self, file_path, sheet_name=None):
        """
        Analyze the structure of an Excel file to determine column layouts and patterns.
        
        Args:
            file_path: Path to the Excel file
            sheet_name: Name of the sheet to analyze
            
        Returns:
            dict: File structure analysis result
        """
        try:
            logging.info(f"Analyzing file structure: {file_path}, sheet: {sheet_name}")
            
            # Try to read the file
            if sheet_name:
                df = pd.read_excel(file_path, sheet_name=sheet_name)
            else:
                # If no sheet specified, get first sheet
                sheets = self.get_sheet_names(file_path)
                if not sheets:
                    raise ValueError("No sheets found in Excel file")
                sheet_name = sheets[0]
                df = pd.read_excel(file_path, sheet_name=sheet_name)
            
            # Get basic info
            row_count = len(df)
            col_count = len(df.columns)
            logging.info(f"File has {row_count} rows and {col_count} columns")
            
            # Analyze column structure
            column_data = {}
            for col in df.columns:
                # Get data type and sample values
                non_null_values = df[col].dropna()
                if len(non_null_values) == 0:
                    data_type = "unknown"
                    distinct_values = []
                    sample_values = []
                else:
                    # Determine general data type category
                    if pd.api.types.is_numeric_dtype(non_null_values):
                        if all(non_null_values.apply(lambda x: int(x) == x)):
                            data_type = "integer"
                        else:
                            data_type = "float"
                    elif pd.api.types.is_datetime64_dtype(non_null_values):
                        data_type = "datetime"
                    else:
                        data_type = "string"
                    
                    # Get distinct and sample values
                    distinct_values = non_null_values.unique().tolist()
                    if len(distinct_values) > 10:
                        distinct_values = distinct_values[:10]  # Limit to 10 distinct values
                    
                    sample_values = non_null_values.head(5).tolist()
                
                # Store column data
                column_data[col] = {
                    "data_type": data_type,
                    "non_null_count": len(non_null_values),
                    "null_count": len(df) - len(non_null_values),
                    "distinct_values": distinct_values[:10] if isinstance(distinct_values, list) else [],
                    "sample_values": sample_values
                }
            
            # Try to identify specific columns based on content
            column_mapping_suggestions = self._suggest_column_mappings(df, column_data)
            
            # Determine if file contains deductible data and identify pattern
            deductible_pattern = self._identify_deductible_pattern(df, column_data)
            
            # Analyze data patterns
            patterns = {
                "has_deductible_data": deductible_pattern["has_deductible_data"],
                "deductible_pattern": deductible_pattern["pattern"],
                "deductible_values": deductible_pattern["values"]
            }
            
            structure = {
                "file_path": file_path,
                "sheet_name": sheet_name,
                "row_count": row_count,
                "column_count": col_count,
                "columns": column_data,
                "column_mapping_suggestions": column_mapping_suggestions,
                "patterns": patterns
            }
            
            logging.info(f"File analysis complete. Identified {len(column_mapping_suggestions)} potential column mappings")
            return structure
            
        except Exception as e:
            error_msg = f"Error analyzing file structure: {str(e)}"
            logging.error(error_msg, exc_info=True)
            raise ValueError(error_msg)
    
    def _suggest_column_mappings(self, df, column_data):
        """
        Suggest mappings between source columns and standard fields.
        
        Args:
            df: DataFrame to analyze
            column_data: Column data from structure analysis
            
        Returns:
            dict: Mapping suggestions with confidence scores
        """
        # Standard fields we want to map to
        standard_fields = {
            "coverage": ["coverage", "coveragename", "coverage name", "coverage type", "cov", "plan", "product"],
            "term": ["term", "termmonths", "term months", "months", "term length"],
            "miles": ["miles", "termmiles", "term miles", "mileage"],
            "frommiles": ["frommiles", "from miles", "min miles", "minmiles", "miles from", "starting miles"],
            "tomiles": ["tomiles", "to miles", "max miles", "maxmiles", "miles to", "ending miles"],
            "minyear": ["minyear", "min year", "year min", "model year min", "minimum year"],
            "maxyear": ["maxyear", "max year", "year max", "model year max", "maximum year"],
            "class": ["class", "vehicle class", "rate class", "rateclass", "class code"],
            "ratecost": ["ratecost", "rate", "cost", "price", "premium", "dealer cost", "dealercost"],
            "deductible": ["deductible", "ded", "deduct", "deductable"]
        }
        
        # Additional patterns to match (regex)
        pattern_matchers = {
            "coverage": [r"(?i)coverage", r"(?i)plan", r"(?i)product"],
            "term": [r"(?i)term", r"(?i)month"],
            "miles": [r"(?i)mile"],
            "class": [r"(?i)class", r"(?i)tier"],
            "ratecost": [r"(?i)rate", r"(?i)cost", r"(?i)price"],
            "deductible": [r"(?i)deduct"]
        }
        
        # Result dictionary
        suggestions = {}
        
        # For each standard field, find potential matches
        for std_field, keywords in standard_fields.items():
            best_match = None
            best_score = 0
            
            # Check each column
            for col in df.columns:
                col_lower = str(col).lower()
                score = 0
                
                # Exact match gets highest score
                if col_lower == std_field:
                    score = 100
                # Exact match to any keyword
                elif col_lower in keywords:
                    score = 90
                # Contains exact keyword
                elif any(keyword in col_lower for keyword in keywords):
                    score = 80
                # Keyword contains column name (for short column names)
                elif len(col_lower) >= 3 and any(col_lower in keyword for keyword in keywords):
                    score = 70
                # Regex pattern matching
                elif std_field in pattern_matchers and any(re.search(pattern, col) for pattern in pattern_matchers[std_field]):
                    score = 60
                
                # Check column contents for additional clues
                if score > 0 and std_field in column_data:
                    col_data = column_data[col]
                    
                    # Check data type appropriateness
                    if std_field in ["term", "miles", "frommiles", "tomiles", "minyear", "maxyear", "class"] and col_data["data_type"] in ["integer", "float"]:
                        score += 5
                    elif std_field in ["ratecost"] and col_data["data_type"] in ["float", "integer"]:
                        score += 5
                    elif std_field in ["deductible"] and col_data["data_type"] in ["integer", "float"]:
                        score += 5
                    elif std_field in ["coverage"] and col_data["data_type"] == "string":
                        score += 5
                
                # Update best match if better score
                if score > best_score:
                    best_score = score
                    best_match = col
            
            # If we have a match with reasonable confidence, add to suggestions
            if best_match and best_score >= 50:
                suggestions[std_field] = {
                    "suggested_column": best_match,
                    "confidence": best_score
                }
        
        # Look for common patterns in data to improve matches
        
        # 1. Check for deductible patterns in column data
        deduct_pattern = self._identify_deductible_column(df, column_data)
        if deduct_pattern["found"] and "deductible" not in suggestions:
            suggestions["deductible"] = {
                "suggested_column": deduct_pattern["column"],
                "confidence": deduct_pattern["confidence"]
            }
        
        # 2. Look for rate class patterns (usually numeric or letter-based classes)
        class_pattern = self._identify_class_column(df, column_data)
        if class_pattern["found"] and "class" not in suggestions:
            suggestions["class"] = {
                "suggested_column": class_pattern["column"],
                "confidence": class_pattern["confidence"]
            }
        
        # Log suggestions
        for field, data in suggestions.items():
            logging.info(f"Suggested mapping: {field} -> {data['suggested_column']} (confidence: {data['confidence']})")
        
        return suggestions
    
    def _identify_deductible_column(self, df, column_data):
        """
        Identify a column that likely contains deductible values.
        
        Args:
            df: DataFrame to analyze
            column_data: Column data from structure analysis
            
        Returns:
            dict: Result with found status, column name, and confidence
        """
        result = {"found": False, "column": None, "confidence": 0}
        
        # Common deductible values
        common_deductibles = [0, 50, 100, 200, 250, 500, 1000]
        
        # Check each numeric column
        for col, data in column_data.items():
            if data["data_type"] in ["integer", "float"]:
                # Get distinct values
                distinct_values = df[col].dropna().unique()
                
                # Calculate what percentage of values are common deductibles
                common_count = sum(1 for val in distinct_values if val in common_deductibles)
                if len(distinct_values) > 0:
                    common_ratio = common_count / len(distinct_values)
                    
                    # If most values are common deductibles, this is likely a deductible column
                    if common_ratio >= 0.5 and common_count >= 2:
                        confidence = int(common_ratio * 100)
                        
                        # If column name contains "deduct", increase confidence
                        if "deduct" in str(col).lower():
                            confidence += 20
                            
                        # Update if better than current
                        if confidence > result["confidence"]:
                            result = {"found": True, "column": col, "confidence": min(confidence, 100)}
        
        return result
    
    def _identify_class_column(self, df, column_data):
        """
        Identify a column that likely contains rate class values.
        
        Args:
            df: DataFrame to analyze
            column_data: Column data from structure analysis
            
        Returns:
            dict: Result with found status, column name, and confidence
        """
        result = {"found": False, "column": None, "confidence": 0}
        
        # Check each column
        for col, data in column_data.items():
            confidence = 0
            
            # Class columns often have few distinct values
            if data["non_null_count"] > 0:
                distinct_count = len(data["distinct_values"])
                
                # Classes usually have between 2-10 values
                if 2 <= distinct_count <= 10:
                    confidence += 30
                    
                    # Class values are usually short (1-3 characters)
                    if all(len(str(val)) <= 3 for val in data["distinct_values"]):
                        confidence += 20
                    
                    # Class columns often contain numbers 1-9 or letters A-F
                    if all(str(val).isdigit() or str(val).upper() in "ABCDEF" for val in data["distinct_values"]):
                        confidence += 20
                    
                    # If column name contains "class", increase confidence
                    if "class" in str(col).lower():
                        confidence += 30
                    
                    # Update if better than current
                    if confidence > result["confidence"]:
                        result = {"found": True, "column": col, "confidence": min(confidence, 100)}
        
        return result
    
    def _identify_deductible_pattern(self, df, column_data):
        """
        Identify how deductible data is structured in the file.
        
        Args:
            df: DataFrame to analyze
            column_data: Column data from structure analysis
            
        Returns:
            dict: Result with pattern information
        """
        result = {
            "has_deductible_data": False,
            "pattern": "unknown",
            "values": []
        }
        
        # Check if there's a deductible column
        deduct_col = None
        for col, data in column_data.items():
            if "deduct" in str(col).lower():
                deduct_col = col
                break
        
        # If no explicit deductible column, look for one using heuristics
        if not deduct_col:
            deduct_pattern = self._identify_deductible_column(df, column_data)
            if deduct_pattern["found"]:
                deduct_col = deduct_pattern["column"]
        
        # If we found a deductible column
        if deduct_col:
            result["has_deductible_data"] = True
            result["pattern"] = "single_column"
            
            # Get unique deductible values
            deduct_values = sorted(df[deduct_col].dropna().unique().tolist())
            result["values"] = deduct_values
        
        # Check for separate deductible columns (like Deduct0, Deduct50, etc.)
        deduct_columns = [col for col in df.columns if re.match(r'(?i)deduct(?:ible)?[\s_]?\d+$', str(col))]
        if deduct_columns:
            result["has_deductible_data"] = True
            result["pattern"] = "multiple_columns"
            
            # Extract the deductible values from column names
            deduct_values = []
            for col in deduct_columns:
                match = re.search(r'(\d+)$', str(col))
                if match:
                    deduct_values.append(int(match.group(1)))
            
            result["values"] = sorted(deduct_values)
        
        return result
    
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
                elif any(term in col_lower for term in ['company', 'co', 'carrier']):
                    if any(term in col_lower for term in ['code', 'id', 'number']):
                        col_info['possible_type'] = 'companycode'
                elif any(term in col_lower for term in ['state', 'province', 'region']):
                    col_info['possible_type'] = 'state'
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
            'companycode': ['companycode', 'company_code', 'company', 'code'],
            'term': ['term', 'termmonths', 'termlength', 'months', 'duration'],
            'miles': ['miles', 'mileage', 'odometer', 'distance', 'milelimit'],
            'frommiles': ['frommiles', 'from_miles', 'startmiles', 'minmiles', 'lowmiles'],
            'tomiles': ['tomiles', 'to_miles', 'endmiles', 'maxmiles', 'highmiles'],
            'coverage': ['coverage', 'coveragetype', 'cover', 'protection', 'plan'],
            'state': ['state', 'location', 'region', 'territory'],
            'class': ['class', 'vehicleclass', 'category', 'tier', 'classification'],
            'plandeduct': ['plandeduct', 'plan_deduct', 'plandeductible', 'plan_deductible'],
            'deduct0': ['deduct0', 'deductible0', 'deductible_0'],
            'deduct50': ['deduct50', 'deductible50', 'deductible_50'],
            'deduct100': ['deduct100', 'deductible100', 'deductible_100'],
            'deduct200': ['deduct200', 'deductible200', 'deductible_200'],
            'deduct250': ['deduct250', 'deductible250', 'deductible_250'],
            'deduct500': ['deduct500', 'deductible500', 'deductible_500'],
            'markup': ['markup', 'mark_up', 'margin', 'profit'],
            'new/used': ['new/used', 'newused', 'condition', 'vehicle_condition'],
            'maxyears': ['maxyears', 'max_years', 'yearmax', 'endyear', 'toyear'],
            'surchargecode': ['surchargecode', 'surcharge_code', 'surcharge'],
            'plancode': ['plancode', 'plan_code', 'plan'],
            'ratecardcode': ['ratecardcode', 'rate_card_code', 'ratecode'],
            'classlistcode': ['classlistcode', 'class_list_code', 'classlist'],
            'minyear': ['minyear', 'min_year', 'yearmin', 'startyear', 'fromyear'],
            'incsccode': ['incsccode', 'inc_sc_code', 'incsc_code'],
            'incscamt': ['incscamt', 'inc_sc_amt', 'incsc_amt']
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
            'companycode': ['companycode', 'company_code', 'company', 'code'],
            'term': ['term', 'termmonths', 'termlength', 'months', 'duration'],
            'miles': ['miles', 'mileage', 'odometer', 'distance', 'milelimit'],
            'frommiles': ['frommiles', 'from_miles', 'startmiles', 'minmiles', 'lowmiles'],
            'tomiles': ['tomiles', 'to_miles', 'endmiles', 'maxmiles', 'highmiles'],
            'coverage': ['coverage', 'coveragetype', 'cover', 'protection', 'plan'],
            'state': ['state', 'location', 'region', 'territory'],
            'class': ['class', 'vehicleclass', 'category', 'tier', 'classification'],
            'plandeduct': ['plandeduct', 'plan_deduct', 'plandeductible', 'plan_deductible'],
            'deduct0': ['deduct0', 'deductible0', 'deductible_0'],
            'deduct50': ['deduct50', 'deductible50', 'deductible_50'],
            'deduct100': ['deduct100', 'deductible100', 'deductible_100'],
            'deduct200': ['deduct200', 'deductible200', 'deductible_200'],
            'deduct250': ['deduct250', 'deductible250', 'deductible_250'],
            'deduct500': ['deduct500', 'deductible500', 'deductible_500'],
            'markup': ['markup', 'mark_up', 'margin', 'profit'],
            'new/used': ['new/used', 'newused', 'condition', 'vehicle_condition'],
            'maxyears': ['maxyears', 'max_years', 'yearmax', 'endyear', 'toyear'],
            'surchargecode': ['surchargecode', 'surcharge_code', 'surcharge'],
            'plancode': ['plancode', 'plan_code', 'plan'],
            'ratecardcode': ['ratecardcode', 'rate_card_code', 'ratecode'],
            'classlistcode': ['classlistcode', 'class_list_code', 'classlist'],
            'minyear': ['minyear', 'min_year', 'yearmin', 'startyear', 'fromyear'],
            'incsccode': ['incsccode', 'inc_sc_code', 'incsc_code'],
            'incscamt': ['incscamt', 'inc_sc_amt', 'incsc_amt']
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