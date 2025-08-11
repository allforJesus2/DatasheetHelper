#!/usr/bin/env python3
"""
Dynamic Relationship System - Core system for managing relationships between data types
"""

import json
import os
from typing import Dict, Any, Optional, List


class DynamicRelationshipSystem:
    """Core system for managing dynamic relationships between data types"""
    
    def __init__(self):
        self.data_types = {}
        self.relationships = {}  # Maps data_type -> {target_data_type: relationship_config}
        self.primary_data_type = None  # Can be set to any data type
        self.data = {}  # Stores actual data for each data type
        
    def add_data_type(self, data_type: str, config: Dict[str, Any]) -> bool:
        """Add a new data type with configuration"""
        if data_type in self.data_types:
            print(f"Data type '{data_type}' already exists")
            return False
            
        self.data_types[data_type] = config
        self.relationships[data_type] = {}
        self.data[data_type] = {}
        print(f"Added data type: {data_type}")
        return True
        
    def remove_data_type(self, data_type: str) -> bool:
        """Remove a data type and all its relationships"""
        if data_type not in self.data_types:
            print(f"Data type '{data_type}' not found")
            return False
            
        # Remove from data types
        del self.data_types[data_type]
        del self.data[data_type]
        
        # Remove relationships where this data type is the source
        if data_type in self.relationships:
            del self.relationships[data_type]
            
        # Remove relationships where this data type is the target
        for source_type in list(self.relationships.keys()):
            if data_type in self.relationships[source_type]:
                del self.relationships[source_type][data_type]
                
        # Update primary data type if needed
        if self.primary_data_type == data_type:
            self.primary_data_type = None
            
        print(f"Removed data type: {data_type}")
        return True
        
    def set_primary_data_type(self, data_type: str) -> bool:
        """Set which data type is the primary source"""
        if data_type not in self.data_types:
            print(f"Data type '{data_type}' not found")
            return False
            
        self.primary_data_type = data_type
        print(f"Primary data type set to: {data_type}")
        return True
        
    def get_primary_data_type(self) -> Optional[str]:
        """Get the current primary data type"""
        return self.primary_data_type
        
    def add_relationship(self, source_data_type: str, target_data_type: str, 
                        relationship_config: Dict[str, Any]) -> bool:
        """Add a relationship between two data types
        
        relationship_config should contain:
        - source_key: The key in source data type to use for lookup
        - target_key: The key in target data type to match against
        - transformation: Optional transformation code (Python expression)
        - description: Optional description of the relationship
        """
        if source_data_type not in self.data_types:
            print(f"Source data type '{source_data_type}' not found")
            return False
            
        if target_data_type not in self.data_types:
            print(f"Target data type '{target_data_type}' not found")
            return False
            
        if source_data_type not in self.relationships:
            self.relationships[source_data_type] = {}
            
        self.relationships[source_data_type][target_data_type] = relationship_config
        print(f"Added relationship: {source_data_type} -> {target_data_type}")
        return True
        
    def remove_relationship(self, source_data_type: str, target_data_type: str) -> bool:
        """Remove a relationship between two data types"""
        if (source_data_type in self.relationships and 
            target_data_type in self.relationships[source_data_type]):
            del self.relationships[source_data_type][target_data_type]
            print(f"Removed relationship: {source_data_type} -> {target_data_type}")
            return True
        else:
            print(f"Relationship not found: {source_data_type} -> {target_data_type}")
            return False
            
    def get_relationships(self, data_type: str = None) -> Dict[str, Any]:
        """Get all relationships or relationships for a specific data type"""
        if data_type is None:
            return self.relationships
        elif data_type in self.relationships:
            return self.relationships[data_type]
        else:
            return {}
            
    def set_data(self, data_type: str, data: Dict[str, Any]) -> bool:
        """Set data for a specific data type"""
        if data_type not in self.data_types:
            print(f"Data type '{data_type}' not found")
            return False
            
        self.data[data_type] = data
        print(f"Set data for {data_type}: {len(data)} items")
        return True
        
    def get_data(self, data_type: str) -> Dict[str, Any]:
        """Get data for a specific data type"""
        return self.data.get(data_type, {})
        
    def get_all_data(self) -> Dict[str, Dict[str, Any]]:
        """Get all data"""
        return self.data.copy()
        
    def process_coordinates_dynamic(self) -> Dict[str, Dict[str, Any]]:
        """Process coordinates using dynamic relationships"""
        if not self.primary_data_type:
            print("No primary data type set")
            return {}
            
        primary_data = self.data.get(self.primary_data_type, {})
        if not primary_data:
            print(f"No data available for primary data type: {self.primary_data_type}")
            return {}
            
        results = {}
        
        # Process each item in the primary data type
        for item_key, item_data in primary_data.items():
            if not item_key:
                continue
                
            # Initialize coordinate data for this item
            coordinate_data = {}
            
            # Process coordinates for the primary data type
            primary_coords = self.data_types[self.primary_data_type].get('coordinate_values', {})
            for coord, field in primary_coords.items():
                coordinate_data[coord] = item_data.get(field)
            
            # Process relationships to other data types
            if self.primary_data_type in self.relationships:
                for target_data_type, relationship_config in self.relationships[self.primary_data_type].items():
                    target_data = self.data.get(target_data_type, {})
                    if not target_data:
                        continue
                        
                    # Get the lookup value from primary data
                    source_key = relationship_config.get('source_key')
                    if not source_key or source_key not in item_data:
                        continue
                        
                    lookup_value = item_data[source_key]
                    
                    # Apply transformation if specified
                    transformation = relationship_config.get('transformation')
                    if transformation:
                        try:
                            # Create a safe environment for eval
                            safe_dict = {'x': lookup_value}
                            lookup_value = eval(transformation, {"__builtins__": {}}, safe_dict)
                        except Exception as e:
                            print(f"Transformation error for {item_key}: {e}")
                            continue
                    
                    # Find matching item in target data
                    target_key = relationship_config.get('target_key')
                    matching_item = None
                    
                    for target_item_key, target_item_data in target_data.items():
                        if target_key in target_item_data and target_item_data[target_key] == lookup_value:
                            matching_item = target_item_data
                            break
                    
                    if matching_item:
                        # Process coordinates for the target data type
                        target_coords = self.data_types[target_data_type].get('coordinate_values', {})
                        for coord, field in target_coords.items():
                            coordinate_data[coord] = matching_item.get(field)
            
            results[item_key] = coordinate_data
        
        return results
        
    def get_data_type_config(self, data_type: str) -> Optional[Dict[str, Any]]:
        """Get configuration for a specific data type"""
        return self.data_types.get(data_type)
        
    def get_all_data_types(self) -> List[str]:
        """Get list of all data type keys"""
        return list(self.data_types.keys())
        
    def get_data_types_dict(self) -> Dict[str, Dict[str, Any]]:
        """Get the full data types dictionary with configurations"""
        return self.data_types.copy()
        
    def get_data_type_names(self) -> Dict[str, str]:
        """Get mapping of data type keys to display names"""
        return {key: config.get('name', key.upper()) for key, config in self.data_types.items()}
        
    def update_data_type_config(self, data_type: str, config: Dict[str, Any]):
        """Update the configuration for a data type"""
        if data_type in self.data_types:
            self.data_types[data_type].update(config)
        else:
            raise ValueError(f"Data type '{data_type}' does not exist")
        
    def save_configuration(self, filepath: str) -> bool:
        """Save the current configuration to a JSON file"""
        try:
            config = {
                'data_types': self.data_types,
                'relationships': self.relationships,
                'primary_data_type': self.primary_data_type
            }
            
            with open(filepath, 'w') as f:
                json.dump(config, f, indent=2)
                
            print(f"Configuration saved to: {filepath}")
            return True
        except Exception as e:
            print(f"Error saving configuration: {e}")
            return False
            
    def load_configuration(self, filepath: str) -> bool:
        """Load configuration from a JSON file"""
        try:
            with open(filepath, 'r') as f:
                config = json.load(f)
                
            self.data_types = config.get('data_types', {})
            self.relationships = config.get('relationships', {})
            self.primary_data_type = config.get('primary_data_type')
            
            # Initialize data storage for loaded data types
            for data_type in self.data_types:
                if data_type not in self.data:
                    self.data[data_type] = {}
                    
            print(f"Configuration loaded from: {filepath}")
            return True
        except Exception as e:
            print(f"Error loading configuration: {e}")
            return False
            
    def validate_system(self) -> List[str]:
        """Validate the system configuration and return any issues"""
        issues = []
        
        # Check if primary data type exists
        if self.primary_data_type and self.primary_data_type not in self.data_types:
            issues.append(f"Primary data type '{self.primary_data_type}' not found in data types")
            
        # Check relationships reference valid data types
        for source_type, targets in self.relationships.items():
            if source_type not in self.data_types:
                issues.append(f"Relationship source '{source_type}' not found in data types")
                
            for target_type in targets:
                if target_type not in self.data_types:
                    issues.append(f"Relationship target '{target_type}' not found in data types")
                    
        # Check coordinate mappings reference valid fields
        for data_type, config in self.data_types.items():
            coordinate_values = config.get('coordinate_values', {})
            headers = config.get('default_headers', [])
            
            for coord, field in coordinate_values.items():
                if field not in headers:
                    issues.append(f"Coordinate '{coord}' in '{data_type}' references field '{field}' not in headers")
                    
        return issues 