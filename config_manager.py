#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Configuration Manager module for Moxy Rates Template Transfer

This module provides classes for managing application settings and 
saved column mappings.
"""

import os
import json
import logging
import configparser
from datetime import datetime


class ConfigManager:
    """Manages application configuration settings."""
    
    def __init__(self, config_file=None):
        """
        Initialize the configuration manager.
        
        Args:
            config_file: Path to the configuration file (optional)
        """
        # Default config path is in the same directory as the script
        if config_file is None:
            self.config_file = os.path.join(
                os.path.dirname(os.path.abspath(__file__)),
                "config.ini"
            )
        else:
            self.config_file = config_file
            
        self.config = configparser.ConfigParser()
        self.load_config()
        
        logging.info(f"ConfigManager initialized with config file: {self.config_file}")
    
    def load_config(self):
        """Load configuration from file."""
        if os.path.exists(self.config_file):
            try:
                self.config.read(self.config_file)
                logging.info("Configuration loaded successfully")
            except Exception as e:
                logging.error(f"Error loading configuration: {str(e)}")
                # Initialize with default sections
                self._init_default_config()
        else:
            logging.info("Configuration file not found, creating default")
            self._init_default_config()
    
    def _init_default_config(self):
        """Initialize default configuration sections."""
        if 'Directories' not in self.config:
            self.config['Directories'] = {}
            
        if 'Options' not in self.config:
            self.config['Options'] = {
                'remember_directories': 'True',
                'open_after_processing': 'True',
                'enable_logging': 'False',
                'auto_detect_formats': 'True',
                'use_saved_mappings': 'True',
                'adjusted_sheet_name': 'Dealer Cost Rates',
                'template_sheet_name': 'Sheet1'
            }
            
        if 'Advanced' not in self.config:
            self.config['Advanced'] = {
                'mapping_confidence_threshold': '70',
                'fuzzy_match_threshold': '60'
            }
    
    def save_config(self):
        """Save configuration to file."""
        try:
            # Ensure directory exists
            os.makedirs(os.path.dirname(self.config_file), exist_ok=True)
            
            with open(self.config_file, 'w') as f:
                self.config.write(f)
                
            logging.info("Configuration saved successfully")
        except Exception as e:
            logging.error(f"Error saving configuration: {str(e)}")
    
    def get_setting(self, key, default=None, section='Options'):
        """
        Get a configuration setting.
        
        Args:
            key: Setting key
            default: Default value if not found
            section: Configuration section
            
        Returns:
            Setting value with appropriate type conversion
        """
        if section not in self.config:
            return default
            
        if key not in self.config[section]:
            return default
            
        value = self.config[section][key]
        
        # Convert to appropriate type based on default
        if isinstance(default, bool):
            return value.lower() in ('true', 'yes', '1', 'on')
        elif isinstance(default, int):
            try:
                return int(value)
            except (ValueError, TypeError):
                return default
        elif isinstance(default, float):
            try:
                return float(value)
            except (ValueError, TypeError):
                return default
        
        return value
    
    def set_setting(self, key, value, section='Options'):
        """
        Set a configuration setting.
        
        Args:
            key: Setting key
            value: Setting value
            section: Configuration section
        """
        if section not in self.config:
            self.config[section] = {}
            
        # Convert value to string for storage
        if isinstance(value, bool):
            self.config[section][key] = str(value)
        else:
            self.config[section][key] = str(value)


class MappingConfigManager:
    """Manages saving and loading of column mappings."""
    
    def __init__(self, config_path=None):
        """
        Initialize the mapping configuration manager.
        
        Args:
            config_path: Path to the mapping configuration file (optional)
        """
        # Default config path is in the same directory as the script
        if config_path is None:
            self.config_path = os.path.join(
                os.path.dirname(os.path.abspath(__file__)),
                "mappings.json"
            )
        else:
            self.config_path = config_path
            
        self.mappings = self._load_mappings()
        
        logging.info(f"MappingConfigManager initialized with config file: {self.config_path}")
    
    def _load_mappings(self):
        """
        Load saved mappings from file.
        
        Returns:
            dict: Loaded mappings or default structure
        """
        if not os.path.exists(self.config_path):
            logging.info("Mappings file not found, creating default structure")
            return self._create_default_mappings()
            
        try:
            with open(self.config_path, 'r') as f:
                mappings = json.load(f)
                logging.info(f"Loaded {len(mappings.get('file_mappings', {}))} file mappings and {len(mappings.get('named_templates', {}))} named templates")
                return mappings
        except Exception as e:
            logging.error(f"Error loading mappings: {str(e)}")
            return self._create_default_mappings()
    
    def _create_default_mappings(self):
        """
        Create default mappings structure.
        
        Returns:
            dict: Default mappings structure
        """
        return {
            "file_mappings": {},
            "named_templates": {},
            "metadata": {
                "created": datetime.now().isoformat(),
                "version": "1.0"
            }
        }
    
    def _save_mappings(self):
        """Save mappings to file."""
        try:
            # Ensure directory exists
            os.makedirs(os.path.dirname(self.config_path), exist_ok=True)
            
            # Update metadata
            if "metadata" not in self.mappings:
                self.mappings["metadata"] = {}
                
            self.mappings["metadata"]["updated"] = datetime.now().isoformat()
            
            with open(self.config_path, 'w') as f:
                json.dump(self.mappings, f, indent=2)
                
            logging.info("Mappings saved successfully")
        except Exception as e:
            logging.error(f"Error saving mappings: {str(e)}")
    
    def get_saved_mapping(self, file_signature):
        """
        Get a mapping for a specific file signature.
        
        Args:
            file_signature: Unique signature for the file structure
            
        Returns:
            dict: Mapping dictionary or None if not found
        """
        mapping = self.mappings.get("file_mappings", {}).get(file_signature)
        
        if mapping:
            logging.info(f"Found saved mapping for signature: {file_signature}")
            # Update last used timestamp
            if isinstance(mapping, dict) and "metadata" not in mapping:
                mapping["metadata"] = {}
                
            if isinstance(mapping, dict) and isinstance(mapping.get("metadata"), dict):
                mapping["metadata"]["last_used"] = datetime.now().isoformat()
                self._save_mappings()
                
        return mapping
    
    def save_mapping(self, file_signature, mapping, mapping_name=None):
        """
        Save a mapping for future use.
        
        Args:
            file_signature: Unique signature for the file structure
            mapping: Dictionary of column mappings
            mapping_name: Optional template name for this mapping
        """
        # Add metadata to mapping
        mapping_with_meta = dict(mapping)  # Create a copy
        
        # Add metadata if it doesn't exist
        if "metadata" not in mapping_with_meta:
            mapping_with_meta["metadata"] = {}
            
        mapping_with_meta["metadata"].update({
            "saved": datetime.now().isoformat(),
            "last_used": datetime.now().isoformat()
        })
        
        if mapping_name:
            mapping_with_meta["metadata"]["name"] = mapping_name
        
        # Save under file signatures
        if "file_mappings" not in self.mappings:
            self.mappings["file_mappings"] = {}
            
        self.mappings["file_mappings"][file_signature] = mapping_with_meta
        
        # If a template name is provided, save as named template too
        if mapping_name:
            if "named_templates" not in self.mappings:
                self.mappings["named_templates"] = {}
                
            self.mappings["named_templates"][mapping_name] = mapping_with_meta
        
        self._save_mappings()
        logging.info(f"Saved mapping for signature: {file_signature}" + 
                    (f" with name: '{mapping_name}'" if mapping_name else ""))
    
    def get_template_names(self):
        """
        Get list of available named templates.
        
        Returns:
            list: List of template names
        """
        return list(self.mappings.get("named_templates", {}).keys())
    
    def get_template(self, template_name):
        """
        Get a specific named template.
        
        Args:
            template_name: Name of the template
            
        Returns:
            dict: Template mapping or None if not found
        """
        return self.mappings.get("named_templates", {}).get(template_name)
    
    def delete_template(self, template_name):
        """
        Delete a named template.
        
        Args:
            template_name: Name of the template to delete
            
        Returns:
            bool: True if deleted, False if not found
        """
        if template_name in self.mappings.get("named_templates", {}):
            del self.mappings["named_templates"][template_name]
            self._save_mappings()
            logging.info(f"Deleted template: {template_name}")
            return True
        
        logging.warning(f"Template not found for deletion: {template_name}")
        return False
    
    def get_recent_mappings(self, limit=5):
        """
        Get recently used mappings.
        
        Args:
            limit: Maximum number of mappings to return
            
        Returns:
            list: List of recent mappings with metadata
        """
        # Collect all mappings with their last_used timestamp
        all_mappings = []
        
        # Add file mappings
        for signature, mapping in self.mappings.get("file_mappings", {}).items():
            if isinstance(mapping, dict) and isinstance(mapping.get("metadata"), dict):
                last_used = mapping["metadata"].get("last_used")
                if last_used:
                    all_mappings.append({
                        "signature": signature,
                        "mapping": mapping,
                        "last_used": last_used,
                        "name": mapping["metadata"].get("name", f"Mapping {signature[:8]}")
                    })
        
        # Sort by last_used timestamp (most recent first)
        all_mappings.sort(key=lambda x: x["last_used"], reverse=True)
        
        # Return limited number
        return all_mappings[:limit] 