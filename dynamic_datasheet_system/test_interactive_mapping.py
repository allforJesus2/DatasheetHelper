#!/usr/bin/env python3
"""
Test script to demonstrate the interactive coordinate mapping functionality
"""

import tkinter as tk
import sys
import os

# Add the current directory to the path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from dynamic_datasheet_app import DynamicDatasheetApp


def test_interactive_mapping():
    """Test the interactive coordinate mapping functionality"""
    print("=== Testing Interactive Coordinate Mapping ===")
    
    # Create the application
    root = tk.Tk()
    app = DynamicDatasheetApp(root)
    
    print("\n1. Setting up sample data:")
    
    # Add sample TD data
    td_data = {
        'TAG001': {'TAG NUMBER': 'TAG001-ABC-123', 'DESCRIPTION': 'Pump 1', 'LOCATION': 'Area A'},
        'TAG002': {'TAG NUMBER': 'TAG002-DEF-456', 'DESCRIPTION': 'Valve 1', 'LOCATION': 'Area B'},
        'TAG003': {'TAG NUMBER': 'TAG003-GHI-789', 'DESCRIPTION': 'Tank 1', 'LOCATION': 'Area C'}
    }
    app.relationship_system.set_data('td', td_data)
    print(f"   - Added {len(td_data)} TD items")
    
    # Add sample PC data
    pc_data = {
        'LINE1': {'Line No.': '123', 'Pressure': '100', 'Temperature': '25'},
        'LINE2': {'Line No.': '456', 'Pressure': '200', 'Temperature': '30'},
        'LINE3': {'Line No.': '789', 'Pressure': '150', 'Temperature': '28'}
    }
    app.relationship_system.set_data('pc', pc_data)
    print(f"   - Added {len(pc_data)} PC items")
    
    print("\n2. Interactive Coordinate Mapping Features:")
    print("   - Click on Excel cells to get coordinates (A1, B2, etc.)")
    print("   - Select data fields from dropdowns")
    print("   - Map coordinates to data fields")
    print("   - Use Add/Update Datasheets to populate Excel")
    print("   - Save mappings to relationship system")
    
    print("\n3. Available Data Fields:")
    for data_type in app.relationship_system.get_all_data_types():
        config = app.relationship_system.get_data_type_config(data_type)
        name = config.get('name', data_type.upper())
        data = app.relationship_system.get_data(data_type)
        if data:
            # Get fields from first item
            for key, value in data.items():
                if isinstance(value, dict):
                    fields = list(value.keys())
                    print(f"   - {name}: {', '.join(fields)}")
                    break
    
    print("\n4. Instructions:")
    print("   1. Select a destination datasheet in the main tab")
    print("   2. Go to the Coordinates tab")
    print("   3. Click 'Interactive Coordinate Mapper'")
    print("   4. Excel will open with the datasheet")
    print("   5. Click on cells in Excel to get coordinates")
    print("   6. Select data fields from dropdowns")
    print("   7. Click 'Add' to map coordinates to fields")
    print("   8. Use 'Add/Update Datasheets' to populate Excel")
    
    print("\n=== Interactive Mapping Test Setup Complete ===")
    print("The system is ready for interactive coordinate mapping!")
    
    # Show the GUI
    print("\nShowing GUI for 30 seconds...")
    root.after(30000, root.destroy)
    root.mainloop()


def demonstrate_mapping_workflow():
    """Demonstrate the complete mapping workflow"""
    print("=== Interactive Coordinate Mapping Workflow ===")
    
    root = tk.Tk()
    app = DynamicDatasheetApp(root)
    
    print("\nWorkflow Steps:")
    print("1. Load Data Sources:")
    print("   - Browse and load TD data from Excel")
    print("   - Browse and load PC data from Excel")
    print("   - Data appears in Data tab")
    
    print("\n2. Configure Relationships:")
    print("   - Go to Relationships tab")
    print("   - Set primary data type (e.g., TD)")
    print("   - Add relationships between data types")
    
    print("\n3. Set Destination:")
    print("   - Browse for destination datasheet in Main tab")
    print("   - This is where data will be populated")
    
    print("\n4. Interactive Coordinate Mapping:")
    print("   - Go to Coordinates tab")
    print("   - Click 'Interactive Coordinate Mapper'")
    print("   - Excel opens with destination datasheet")
    print("   - Click on cells to get coordinates")
    print("   - Select data fields from dropdowns")
    print("   - Map coordinates to fields")
    
    print("\n5. Populate Datasheets:")
    print("   - Use 'Add/Update Datasheets' button")
    print("   - Data is populated into Excel sheets")
    print("   - Each tag gets its own sheet or row")
    
    print("\n6. Process Coordinates:")
    print("   - Use 'Process Coordinates' to generate results")
    print("   - View coordinate processing results")
    print("   - Export to Excel if needed")
    
    print("\n=== Workflow Demonstration Complete ===")
    
    # Show the GUI
    print("\nShowing GUI for 20 seconds...")
    root.after(20000, root.destroy)
    root.mainloop()


if __name__ == "__main__":
    print("Choose a demonstration:")
    print("1. Test interactive coordinate mapping setup")
    print("2. Demonstrate complete mapping workflow")
    
    choice = input("Enter choice (1 or 2): ").strip()
    
    if choice == "1":
        test_interactive_mapping()
    elif choice == "2":
        demonstrate_mapping_workflow()
    else:
        print("Invalid choice. Running interactive mapping test...")
        test_interactive_mapping() 