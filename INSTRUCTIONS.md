# Instructions for Developing a Moxy Rates Template Transfer App for Windows

## **Overview**

This Windows application automates the process of transferring data from an **Adjusted Rates** spreadsheet into a **Template** spreadsheet, ensuring that all data aligns correctly. The application will:

1. Present a user-friendly GUI interface for file selection and configuration
2. **Detect and analyze** the structure of incoming Excel files automatically
3. Load the data from an **Adjusted Rates** Excel file
4. Load the **Template** Excel file
5. **Apply intelligent mapping** between different column naming conventions
6. Transform and map data according to the required format
7. Ensure **alignment** of data based on `Deductible, Miles, FromMiles, ToMiles, Coverage, and Class`
8. Maintain **MinYear** and **MaxYears** in the transformed data
9. Export the processed data back into the template format
10. **Save successful mappings** for future use with similar files
11. Provide clear feedback and error messages to the user

---

## **Application Architecture**

### 1. Application Components

The application will consist of the following components:

- **GUI Interface**: A user-friendly Windows form application
- **Data Processing Core**: Engine that handles the Excel transformations
- **File Format Analyzer**: System that detects and identifies file structures
- **Intelligent Mapping System**: Handles different column naming conventions and structures
- **Configuration Manager**: Saves and loads user preferences and mapping templates
- **Error Handler**: Manages exceptions and provides user feedback
- **Logger**: Records operations and issues for troubleshooting
- **Visual Mapping Interface**: Allows users to manually map columns when automation fails

### 2. Technology Stack

- **Language**: Python 3.9+ (packaged as Windows executable)
- **GUI Framework**: Tkinter or PyQt5 for the interface
- **Data Processing**: Pandas for Excel manipulation
- **Machine Learning**: Optional scikit-learn for advanced pattern recognition in file formats
- **Packaging**: PyInstaller to create standalone Windows executable
- **Dependencies**: 
  - pandas
  - openpyxl (for Excel operations)
  - tkinter/PyQt5 (for GUI)
  - tqdm (for progress indicators)
  - configparser (for saving settings)
  - fuzzywuzzy (for fuzzy matching of column names)
  - scikit-learn (optional, for ML-based pattern recognition)

---

## **Detailed Component Specifications**

### 1. GUI Interface Design

Create a clean, intuitive interface with these elements:

```
┌─ Moxy Rates Template Transfer ─────────────────────────────────┐
│                                                                │
│  ┌─ Input Files ───────────────────────────────────────────┐   │
│  │ Adjusted Rates File:  [___________________] [Browse...] │   │
│  │ Template File:        [___________________] [Browse...] │   │
│  └──────────────────────────────────────────────────────────┘   │
│                                                                │
│  ┌─ Output ─────────────────────────────────────────────────┐   │
│  │ Output Filename:     [___________________] [Browse...] │   │
│  └──────────────────────────────────────────────────────────┘   │
│                                                                │
│  ┌─ Options ────────────────────────────────────────────────┐   │
│  │ ☑ Remember last used directories                         │   │
│  │ ☑ Open output file after processing                      │   │
│  │ □ Enable detailed logging                                │   │
│  │ ☑ Auto-detect file formats                               │   │
│  │ ☑ Use saved mappings when available                      │   │
│  │                                                          │   │
│  │ Adjusted Rates Sheet: [Dealer Cost Rates ▼]              │   │
│  │ Template Sheet:       [Sheet1 ▼]                         │   │
│  │                                                          │   │
│  │ [Manage Saved Mappings]  [Advanced Options]              │   │
│  └──────────────────────────────────────────────────────────┘   │
│                                                                │
│  [Process Files]  [Preview Mapping]  [Exit]                   │
│                                                                │
│  Status: Ready                                                 │
│  [________________________________________________] 0%         │
│                                                                │
└────────────────────────────────────────────────────────────────┘
```

### 2. File Format Analyzer Implementation

```python
def analyze_file_structure(excel_path, sheet_name=None):
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
        # Load the Excel file
        if sheet_name:
            df = pd.read_excel(excel_path, sheet_name=sheet_name)
        else:
            # Try to analyze all sheets if none specified
            xls = pd.ExcelFile(excel_path)
            sheets_analysis = {}
            
            for sheet in xls.sheet_names:
                sheets_analysis[sheet] = analyze_sheet_structure(pd.read_excel(excel_path, sheet_name=sheet))
            
            return {
                'file_path': excel_path,
                'sheets': sheets_analysis,
                'suggested_main_sheet': identify_main_data_sheet(sheets_analysis)
            }
        
        return analyze_sheet_structure(df)
        
    except Exception as e:
        return {'error': str(e)}

def analyze_sheet_structure(df):
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
    
    # Analyze each column
    for col in df.columns:
        col_info = {
            'data_type': str(df[col].dtype),
            'unique_values': df[col].nunique(),
            'sample_values': df[col].dropna().sample(min(5, len(df[col].dropna()))).tolist() if not df[col].empty else [],
            'null_percentage': (df[col].isna().sum() / len(df)) * 100
        }
        
        # Detect if column is likely a key/identifier
        if col_info['unique_values'] > (len(df) * 0.9) and len(df) > 10:
            structure['potential_key_columns'].append(col)
            
        # Detect if column has numeric values
        if df[col].dtype in ['int64', 'float64']:
            col_info['min'] = df[col].min() if not df[col].empty else None
            col_info['max'] = df[col].max() if not df[col].empty else None
            col_info['mean'] = df[col].mean() if not df[col].empty else None
            
            # Detect if column might be a deductible column
            if 'deduct' in col.lower() or (col_info['min'] >= 0 and col_info['max'] <= 1000):
                col_info['possible_type'] = 'deductible'
                
        # Detect if column might be a year column
        if df[col].dtype in ['int64'] and not df[col].empty:
            if df[col].min() >= 1990 and df[col].max() <= 2050:
                col_info['possible_type'] = 'year'
                
        # Look for common column name patterns
        col_lower = col.lower()
        if any(term in col_lower for term in ['mile', 'distance', 'km']):
            col_info['possible_type'] = 'mileage'
        elif any(term in col_lower for term in ['cover', 'coverage']):
            col_info['possible_type'] = 'coverage'
        elif any(term in col_lower for term in ['class', 'category', 'type']):
            col_info['possible_type'] = 'class'
        elif any(term in col_lower for term in ['term', 'duration', 'period']):
            col_info['possible_type'] = 'term'
        elif any(term in col_lower for term in ['rate', 'cost', 'price']):
            col_info['possible_type'] = 'rate'
            
        structure['columns'][col] = col_info
    
    # Detect if this is likely a template or data sheet
    structure['likely_purpose'] = 'template' if len(df) < 5 else 'data'
    
    # Generate column similarity scores for common required fields
    structure['column_mapping_suggestions'] = suggest_column_mappings(structure['columns'])
    
    return structure
    
def suggest_column_mappings(columns_info):
    """
    Suggest possible mappings for columns based on names and data characteristics.
    
    Args:
        columns_info: Dictionary of column information
        
    Returns:
        dict: Suggested mappings for required fields
    """
    required_fields = {
        'coverage': ['coverage', 'cover', 'cov', 'protection'],
        'term': ['term', 'duration', 'period', 'months'],
        'miles': ['miles', 'mileage', 'distance', 'odometer'],
        'frommiles': ['frommiles', 'from_miles', 'start_miles', 'min_miles'],
        'tomiles': ['tomiles', 'to_miles', 'end_miles', 'max_miles'],
        'minyear': ['minyear', 'min_year', 'from_year', 'start_year'],
        'maxyears': ['maxyears', 'max_years', 'to_year', 'end_year', 'year_limit'],
        'class': ['class', 'category', 'vehicle_class', 'classification'],
        'ratecost': ['ratecost', 'rate', 'cost', 'price', 'premium'],
        'deductible': ['deductible', 'deduct', 'ded', 'deductable']
    }
    
    suggestions = {}
    
    # Use fuzzy matching to find the best match for each required field
    from fuzzywuzzy import fuzz
    
    for required_field, synonyms in required_fields.items():
        best_match = None
        best_score = 0
        
        for col_name in columns_info.keys():
            col_lower = col_name.lower()
            
            # Check exact matches first
            if col_lower == required_field:
                best_match = col_name
                best_score = 100
                break
                
            # Check synonyms
            for synonym in synonyms:
                if synonym in col_lower:
                    score = 80  # Good match based on substring
                    if score > best_score:
                        best_match = col_name
                        best_score = score
            
            # Fuzzy matching if no direct match
            if best_score < 80:
                for synonym in synonyms + [required_field]:
                    score = fuzz.ratio(col_lower, synonym)
                    if score > best_score and score > 60:  # Threshold for fuzzy matching
                        best_match = col_name
                        best_score = score
                        
            # Check for possible_type field match
            if 'possible_type' in columns_info[col_name]:
                if columns_info[col_name]['possible_type'] == required_field:
                    score = 90  # High confidence based on detected data type
                    if score > best_score:
                        best_match = col_name
                        best_score = score
        
        if best_match:
            suggestions[required_field] = {
                'suggested_column': best_match,
                'confidence': best_score
            }
            
    return suggestions
```

### 3. Intelligent Mapping System

```python
class MappingSystem:
    """Handles mapping between different column naming conventions"""
    
    def __init__(self, config_manager):
        self.config_manager = config_manager
        self.current_mapping = {}
        self.mapping_confidence = {}
        self.required_fields = [
            'Coverage', 'Term', 'Miles', 'FromMiles', 'ToMiles', 
            'MinYear', 'MaxYears', 'Class', 'RateCost', 'Deductible'
        ]
        
    def generate_mapping(self, source_structure, target_structure=None):
        """
        Generate mapping between source and standardized columns.
        
        Args:
            source_structure: Source file structure analysis
            target_structure: Optional target structure for direct mapping
            
        Returns:
            dict: Mapping between source and standard columns
        """
        # If we have a saved mapping for this file pattern, use it
        file_signature = self._generate_file_signature(source_structure)
        saved_mapping = self.config_manager.get_saved_mapping(file_signature)
        
        if saved_mapping:
            self.current_mapping = saved_mapping
            self.mapping_confidence = {field: 100 for field in saved_mapping}
            return saved_mapping
            
        # Otherwise, generate mapping based on column analysis
        mapping = {}
        confidence = {}
        
        # Use suggested mappings from file analysis
        suggestions = source_structure['column_mapping_suggestions']
        
        for required_field in self.required_fields:
            field_lower = required_field.lower()
            if field_lower in suggestions:
                mapping[required_field] = suggestions[field_lower]['suggested_column']
                confidence[required_field] = suggestions[field_lower]['confidence']
            else:
                # Fall back to direct name matching
                for col in source_structure['columns']:
                    if col.lower() == field_lower:
                        mapping[required_field] = col
                        confidence[required_field] = 100
                        break
        
        self.current_mapping = mapping
        self.mapping_confidence = confidence
        
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
            raise ValueError("No mapping defined. Call generate_mapping first.")
            
        # Create new dataframe with mapped columns
        mapped_df = pd.DataFrame()
        
        for standard_col, source_col in self.current_mapping.items():
            if source_col in df.columns:
                mapped_df[standard_col] = df[source_col]
                
        # Check for missing required columns
        missing = [col for col in self.required_fields if col not in mapped_df.columns]
        if missing:
            raise ValueError(f"Missing required columns after mapping: {', '.join(missing)}")
            
        return mapped_df
        
    def save_current_mapping(self, source_structure, mapping_name=None):
        """
        Save the current mapping for future use.
        
        Args:
            source_structure: Source file structure for signature generation
            mapping_name: Optional name for this mapping template
        """
        if not self.current_mapping:
            raise ValueError("No mapping defined to save")
            
        file_signature = self._generate_file_signature(source_structure)
        self.config_manager.save_mapping(
            file_signature, 
            self.current_mapping,
            mapping_name=mapping_name
        )
        
    def _generate_file_signature(self, structure):
        """Generate a unique signature for a file structure"""
        # Create a signature based on column names, order, and data types
        columns = sorted(list(structure['columns'].keys()))
        data_types = [structure['columns'][col]['data_type'] for col in columns]
        
        import hashlib
        signature = hashlib.md5(
            str(columns).encode() + str(data_types).encode()
        ).hexdigest()
        
        return signature
```

### 4. Visual Mapping Interface

Add a visual interface for manual column mapping when automatic detection fails:

```python
class MappingDialog:
    """Visual interface for manually mapping columns between files"""
    
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
        
        self.source_columns = source_columns
        self.required_fields = required_fields
        self.mapping = suggested_mapping or {}
        self.result_mapping = None
        
        self._create_ui()
        
    def _create_ui(self):
        """Create the UI components for mapping"""
        # Header frame
        header_frame = tk.Frame(self.dialog)
        header_frame.pack(fill=tk.X, padx=10, pady=10)
        
        tk.Label(header_frame, text="Map Source Columns to Required Fields", 
                font=("Arial", 14)).pack(side=tk.LEFT)
        
        # Instructions
        instr_frame = tk.Frame(self.dialog)
        instr_frame.pack(fill=tk.X, padx=10)
        
        tk.Label(instr_frame, 
                text="Select the source column that corresponds to each required field.",
                font=("Arial", 10)).pack(anchor=tk.W)
        
        # Mapping area with scrolling
        mapping_frame = tk.Frame(self.dialog)
        mapping_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Canvas with scrollbar for mapping entries
        canvas = tk.Canvas(mapping_frame)
        scrollbar = tk.Scrollbar(mapping_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas)
        
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
        
        # Add headers
        tk.Label(scrollable_frame, text="Required Field", font=("Arial", 10, "bold"), 
                width=20).grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        tk.Label(scrollable_frame, text="Source Column", font=("Arial", 10, "bold"), 
                width=30).grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
        tk.Label(scrollable_frame, text="Preview", font=("Arial", 10, "bold"), 
                width=20).grid(row=0, column=2, padx=5, pady=5, sticky=tk.W)
        
        # Add mapping rows
        for i, field in enumerate(self.required_fields, 1):
            # Required field name
            tk.Label(scrollable_frame, text=field, width=20).grid(
                row=i, column=0, padx=5, pady=5, sticky=tk.W)
            
            # Dropdown for source column
            var = tk.StringVar()
            self.mapping_vars[field] = var
            
            # Set initial value if in suggested mapping
            if field in self.mapping:
                var.set(self.mapping[field])
            
            # Create dropdown with source columns
            dropdown = ttk.Combobox(scrollable_frame, textvariable=var, 
                                    values=[""] + self.source_columns, width=30)
            dropdown.grid(row=i, column=1, padx=5, pady=5, sticky=tk.W)
            
            # Preview button
            preview_btn = tk.Button(scrollable_frame, text="Preview Data", width=15,
                                   command=lambda f=field: self._preview_mapping(f))
            preview_btn.grid(row=i, column=2, padx=5, pady=5)
        
        # Bottom button frame
        button_frame = tk.Frame(self.dialog)
        button_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # Auto-map button
        auto_btn = tk.Button(button_frame, text="Auto-Map Remaining", 
                            command=self._auto_map_remaining)
        auto_btn.pack(side=tk.LEFT, padx=5)
        
        # Save as template checkbox
        self.save_template_var = tk.BooleanVar(value=False)
        save_cb = tk.Checkbutton(button_frame, text="Save as mapping template", 
                                variable=self.save_template_var)
        save_cb.pack(side=tk.LEFT, padx=20)
        
        # Template name entry
        self.template_name_var = tk.StringVar()
        tk.Label(button_frame, text="Template name:").pack(side=tk.LEFT, padx=5)
        name_entry = tk.Entry(button_frame, textvariable=self.template_name_var, width=20)
        name_entry.pack(side=tk.LEFT, padx=5)
        
        # Cancel/Apply buttons
        cancel_btn = tk.Button(button_frame, text="Cancel", command=self.dialog.destroy)
        cancel_btn.pack(side=tk.RIGHT, padx=5)
        
        apply_btn = tk.Button(button_frame, text="Apply Mapping", command=self._apply_mapping)
        apply_btn.pack(side=tk.RIGHT, padx=5)
    
    def _preview_mapping(self, field):
        """Show a preview of the selected column data"""
        # Implementation would show sample data from the selected mapping
        pass
        
    def _auto_map_remaining(self):
        """Attempt to automatically map remaining unmapped fields"""
        # Implementation would use the MappingSystem to suggest remaining fields
        pass
        
    def _apply_mapping(self):
        """Apply the mapping and close the dialog"""
        # Get the final mapping from the UI
        self.result_mapping = {
            field: var.get() for field, var in self.mapping_vars.items()
            if var.get() != ""  # Only include mapped fields
        }
        
        # Check if we should save as template
        self.save_as_template = self.save_template_var.get()
        self.template_name = self.template_name_var.get() if self.save_as_template else None
        
        # Close the dialog
        self.dialog.destroy()
        
    def show(self):
        """Show the dialog and return the mapping when closed"""
        # Make dialog modal
        self.dialog.transient(self.dialog.master)
        self.dialog.grab_set()
        self.dialog.master.wait_window(self.dialog)
        
        return self.result_mapping, self.save_as_template, self.template_name
```

### 5. Configuration Manager for Saved Mappings

```python
class MappingConfigManager:
    """Manages saving and loading of column mappings"""
    
    def __init__(self, config_path="mappings.json"):
        """Initialize the mapping configuration manager"""
        self.config_path = config_path
        self.mappings = self._load_mappings()
        
    def _load_mappings(self):
        """Load saved mappings from file"""
        import json
        import os
        
        if not os.path.exists(self.config_path):
            return {
                "file_mappings": {},
                "named_templates": {}
            }
            
        try:
            with open(self.config_path, 'r') as f:
                return json.load(f)
        except:
            return {
                "file_mappings": {},
                "named_templates": {}
            }
            
    def _save_mappings(self):
        """Save mappings to file"""
        import json
        
        with open(self.config_path, 'w') as f:
            json.dump(self.mappings, f, indent=2)
            
    def get_saved_mapping(self, file_signature):
        """
        Get a mapping for a specific file signature.
        
        Args:
            file_signature: Unique signature for the file structure
            
        Returns:
            dict: Mapping dictionary or None if not found
        """
        return self.mappings["file_mappings"].get(file_signature)
        
    def save_mapping(self, file_signature, mapping, mapping_name=None):
        """
        Save a mapping for future use.
        
        Args:
            file_signature: Unique signature for the file structure
            mapping: Dictionary of column mappings
            mapping_name: Optional template name for this mapping
        """
        # Save under file signatures
        self.mappings["file_mappings"][file_signature] = mapping
        
        # If a template name is provided, save as named template too
        if mapping_name:
            self.mappings["named_templates"][mapping_name] = mapping
            
        self._save_mappings()
        
    def get_template_names(self):
        """Get list of available named templates"""
        return list(self.mappings["named_templates"].keys())
        
    def get_template(self, template_name):
        """Get a specific named template"""
        return self.mappings["named_templates"].get(template_name)
        
    def delete_template(self, template_name):
        """Delete a named template"""
        if template_name in self.mappings["named_templates"]:
            del self.mappings["named_templates"][template_name]
            self._save_mappings()
```

---

## **Advanced Application Workflow**

1. **Application Start**:
   - Load saved configurations and mapping templates
   - Initialize GUI components
   - Set up logging based on configuration

2. **File Selection**:
   - User selects Adjusted Rates file
   - User selects Template file
   - Optional: User specifies output location and filename

3. **File Format Analysis**:
   - Analyze selected files to detect structure
   - Identify column names, data types, and patterns
   - Generate file signatures for mapping lookup

4. **Mapping Generation**:
   - Check for saved mappings matching file signatures
   - Apply heuristic approaches to suggest column mappings
   - Calculate confidence levels for suggested mappings
   - Display mapping preview if confidence is low

5. **User Mapping Confirmation (if needed)**:
   - Show visual mapping interface for low-confidence mappings
   - Allow user to modify suggested mappings
   - Provide data previews to verify mapping correctness
   - Option to save successful mapping as template

6. **Data Processing**:
   - Display progress bar
   - Load files with confirmed mapping
   - Extract and clean data
   - Transform and pivot data
   - Compute PlanDeduct and other derived fields
   - Apply numeric precision rules

7. **Output Generation**:
   - Apply final mapping to transform data
   - Merge transformed data with template
   - Save to output file
   - Display confirmation message
   - Optionally open the file automatically

8. **Mapping Management**:
   - Save successful mappings for future use
   - Update mapping confidence based on results
   - Allow user to manage saved mappings

9. **Exception Handling**:
   - Catch and handle exceptions gracefully
   - Display user-friendly error messages
   - Log detailed error information
   - Suggest resolution steps

---

## **Managing Saved Mappings UI**

Add a dialog for managing saved mapping templates:

```
┌─ Manage Mapping Templates ─────────────────────────────────────┐
│                                                                │
│  Available Mapping Templates:                                  │
│  ┌─────────────────────────────────────────────────────────┐   │
│  │ Template Name                   │ Last Used    │ Fields │   │
│  │────────────────────────────────┼──────────────┼────────│   │
│  │ Moxy Standard Format           │ 2023-10-15   │   10   │   │
│  │ Legacy Rates Format            │ 2023-09-20   │   8    │   │
│  │ QuickQuote Format              │ 2023-08-05   │   10   │   │
│  │ Dealer Cost Special            │ 2023-07-12   │   12   │   │
│  └─────────────────────────────────────────────────────────┘   │
│                                                                │
│  [View Details]  [Rename]  [Export]  [Import]  [Delete]        │
│                                                                │
│  [Close]                                                       │
│                                                                │
└────────────────────────────────────────────────────────────────┘
```

---

## **Example Heuristic Column Detection Logic**

```python
def identify_column_purpose(column_name, sample_values):
    """
    Use heuristic approaches to identify the purpose of a column.
    
    Args:
        column_name: Name of the column
        sample_values: Sample values from the column
        
    Returns:
        str: Likely purpose of the column
    """
    # Normalize column name
    name = column_name.lower().replace(' ', '').replace('_', '')
    
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
```

---

## **Testing Strategy**

### 1. Unit Testing

Test individual components:
- File loading functions
- Data transformation logic
- Template integration
- File format analysis
- Column mapping heuristics
- Configuration management

### 2. Integration Testing

Test the entire workflow:
- File selection → format analysis → mapping → processing → output
- Error handling scenarios
- Configuration saving/loading
- Mapping template management

### 3. User Acceptance Testing

Create test cases for:
- Different file formats and structures
- Various column naming conventions
- Edge cases (empty files, missing columns)
- Poorly structured input files
- Error recovery
- Mapping interface usability

---

## **Enhanced Development Process**

### Phase 1: Core Functionality
1. Create data processing module
2. Implement file loading and validation
3. Develop transformation logic
4. Build template integration

### Phase 2: Intelligent Format Handling
1. Develop file format analyzer
2. Implement column mapping system
3. Create heuristic detection algorithms
4. Build mapping storage and retrieval

### Phase 3: User Interface
1. Design the GUI layout
2. Implement file selection dialogs
3. Create visual mapping interface
4. Add mapping management UI
5. Create progress indicators

### Phase 4: Error Handling and Logging
1. Implement exception catching
2. Design user-friendly error messages
3. Set up logging system
4. Add validation checks

### Phase 5: Packaging and Distribution
1. Create executable with PyInstaller
2. Design installer
3. Test on different Windows versions
4. Create documentation
5. Create user guide for mapping system

---

## **Application Enhancement Roadmap**

### Version 1.1
- Add batch processing capability for multiple files
- Include template customization options
- Implement data preview functionality
- Add more advanced heuristic detection algorithms

### Version 1.2
- Add reporting features
- Create comparison tools between versions
- Implement backup and restore functionality
- Add machine learning-based format detection

### Version 2.0
- Support for cloud storage integration
- Add scheduling capabilities
- Create template designer
- Implement format converter for non-standard files

---

This document provides detailed instructions for developing a Windows executable application that performs Excel data transformation for Moxy Rates Template Transfer, with robust file format detection, intelligent mapping systems, and user-friendly interfaces for handling various file formats and column naming conventions.

