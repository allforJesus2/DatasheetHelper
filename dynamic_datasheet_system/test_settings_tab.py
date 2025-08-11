#!/usr/bin/env python3
"""
Test script for the Settings tab functionality
"""

import tkinter as tk
from tkinter import ttk
import sys
import os

# Add the current directory to the path for imports
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

try:
    from dynamic_datasheet_app import DynamicDatasheetApp
    from dynamic_relationship_system import DynamicRelationshipSystem
except ImportError as e:
    print(f"Import error: {e}")
    sys.exit(1)

def test_settings_tab():
    """Test the Settings tab functionality"""
    print("Testing Settings tab functionality...")
    
    # Create root window
    root = tk.Tk()
    root.title("Settings Tab Test")
    root.geometry("800x600")
    
    # Create app
    app = DynamicDatasheetApp(root)
    
    # Test data types display
    print("\n1. Testing data types display in Settings tab...")
    
    # Get data types from relationship system
    data_types = app.relationship_system.get_data_types_dict()
    print(f"   Found {len(data_types)} data types:")
    
    for key, config in data_types.items():
        name = config.get('name', key.upper())
        short_name = config.get('short_name', key.upper())
        description = config.get('description', 'No description')
        print(f"   - {key}: {name} ({short_name}) - {description}")
    
    # Test primary data type
    primary = app.relationship_system.get_primary_data_type()
    print(f"   Primary data type: {primary}")
    
    # Test refresh functionality
    print("\n2. Testing refresh functionality...")
    if hasattr(app, 'refresh_data_types_list'):
        app.refresh_data_types_list()
        print("   ✓ refresh_data_types_list method exists and executed")
    else:
        print("   ✗ refresh_data_types_list method not found")
    
    # Test treeview population
    print("\n3. Testing treeview population...")
    if hasattr(app, 'data_types_tree'):
        items = app.data_types_tree.get_children()
        print(f"   Treeview contains {len(items)} items:")
        
        for item in items:
            values = app.data_types_tree.item(item, 'values')
            print(f"   - {values}")
    else:
        print("   ✗ data_types_tree not found")
    
    print("\n4. Testing configuration and removal methods...")
    
    # Test configure method
    if hasattr(app, 'configure_selected_data_type'):
        print("   ✓ configure_selected_data_type method exists")
    else:
        print("   ✗ configure_selected_data_type method not found")
    
    # Test remove method
    if hasattr(app, 'remove_selected_data_type'):
        print("   ✓ remove_selected_data_type method exists")
    else:
        print("   ✗ remove_selected_data_type method not found")
    
    print("\n5. Testing GUI integration...")
    
    # Switch to settings tab
    app.notebook.select(4)  # Settings tab is the 5th tab (index 4)
    print("   ✓ Switched to Settings tab")
    
    print("\nTest completed! Check the GUI to see the Settings tab in action.")
    print("You should see:")
    print("- A treeview showing existing data types (TD, PC)")
    print("- Columns: Key, Name, Short Name, Description, Status")
    print("- Buttons: Configure Selected, Remove Selected, Refresh List")
    print("- The primary data type marked as 'Primary'")
    
    # Run the GUI
    root.mainloop()

def demonstrate_settings_features():
    """Demonstrate the Settings tab features"""
    print("\n" + "="*60)
    print("SETTINGS TAB FEATURES DEMONSTRATION")
    print("="*60)
    
    print("\n1. EXISTING DATA TYPES DISPLAY:")
    print("   - Shows all configured data types in a treeview")
    print("   - Columns: Key, Name, Short Name, Description, Status")
    print("   - Status shows Primary/Secondary and Loaded/Not Loaded")
    
    print("\n2. DATA TYPE MANAGEMENT:")
    print("   - Configure Selected: Opens configuration dialog for selected data type")
    print("   - Remove Selected: Removes data type with confirmation")
    print("   - Refresh List: Updates the display after changes")
    
    print("\n3. INTEGRATION:")
    print("   - Automatically refreshes when data types are added/removed")
    print("   - Integrates with the main GUI refresh system")
    print("   - Shows real-time status of data loading")
    
    print("\n4. DEFAULT DATA TYPES:")
    print("   - TD (Tag Dictionary): Primary data type")
    print("   - PC (Process Conditions): Secondary data type")
    print("   - Both have sample configurations and relationships")

if __name__ == "__main__":
    demonstrate_settings_features()
    test_settings_tab() 