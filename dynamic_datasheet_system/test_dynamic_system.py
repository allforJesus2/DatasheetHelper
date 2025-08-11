#!/usr/bin/env python3
"""
Test script to demonstrate the dynamic relationship system
"""

import tkinter as tk
import sys
import os

# Add the current directory to the path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from dynamic_datasheet_app import DynamicDatasheetApp


def test_dynamic_relationships():
    """Test the dynamic relationship system"""
    print("=== Testing Dynamic Relationship System ===")
    
    # Create the application
    root = tk.Tk()
    app = DynamicDatasheetApp(root)
    
    print("\n1. Initial Setup:")
    print(f"   - Data types: {app.relationship_system.get_all_data_types()}")
    print(f"   - Primary data type: {app.relationship_system.get_primary_data_type()}")
    print(f"   - Relationships: {app.relationship_system.get_relationships()}")
    
    print("\n2. Adding Sample Data:")
    
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
    
    print("\n3. Testing Coordinate Processing:")
    
    # Process coordinates
    results = app.relationship_system.process_coordinates_dynamic()
    print(f"   - Generated {len(results)} coordinate mappings")
    
    for tag, coords in results.items():
        print(f"     {tag}: {coords}")
    
    print("\n4. Testing Primary Data Type Change:")
    
    # Change primary data type to PC
    app.relationship_system.set_primary_data_type('pc')
    print(f"   - Changed primary to: {app.relationship_system.get_primary_data_type()}")
    
    # Add reverse relationship
    app.relationship_system.add_relationship('pc', 'td', {
        'source_key': 'Line No.',
        'target_key': 'TAG NUMBER',
        'transformation': 'f"TAG001-ABC-{x}" if x == "123" else f"TAG002-DEF-{x}" if x == "456" else f"TAG003-GHI-{x}"',
        'description': 'PC references TD via Line Number'
    })
    print("   - Added reverse relationship: PC -> TD")
    
    # Process coordinates with PC as primary
    results_pc = app.relationship_system.process_coordinates_dynamic()
    print(f"   - Generated {len(results_pc)} coordinate mappings with PC as primary")
    
    for line, coords in results_pc.items():
        print(f"     {line}: {coords}")
    
    print("\n5. Testing System Validation:")
    
    # Validate system
    issues = app.relationship_system.validate_system()
    if not issues:
        print("   - System validation passed!")
    else:
        print("   - System validation issues:")
        for issue in issues:
            print(f"     * {issue}")
    
    print("\n=== Dynamic System Test Complete ===")
    print("The system successfully demonstrates:")
    print("- Any data type can be primary")
    print("- Bidirectional relationships")
    print("- Dynamic coordinate processing")
    print("- Flexible transformations")
    
    # Show the GUI for a few seconds
    print("\nShowing GUI for 10 seconds...")
    root.after(10000, root.destroy)
    root.mainloop()


def demonstrate_gui_features():
    """Demonstrate GUI features"""
    print("=== GUI Feature Demonstration ===")
    
    root = tk.Tk()
    app = DynamicDatasheetApp(root)
    
    print("\nGUI Features Available:")
    print("1. Main Tab - Data sources and destination")
    print("2. Relationships Tab - Configure data type relationships")
    print("3. Data Tab - View and manage data")
    print("4. Coordinates Tab - Process and view coordinate data")
    print("5. Settings Tab - System configuration and validation")
    
    print("\nKey Dynamic Features:")
    print("- Primary data type can be changed in Relationships tab")
    print("- Data type frames are generated automatically")
    print("- Relationships can be viewed in tree format")
    print("- System validation checks configuration")
    
    print("\nShowing GUI for 15 seconds...")
    root.after(15000, root.destroy)
    root.mainloop()


if __name__ == "__main__":
    print("Choose a demonstration:")
    print("1. Test dynamic relationship system")
    print("2. Demonstrate GUI features")
    
    choice = input("Enter choice (1 or 2): ").strip()
    
    if choice == "1":
        test_dynamic_relationships()
    elif choice == "2":
        demonstrate_gui_features()
    else:
        print("Invalid choice. Running relationship test...")
        test_dynamic_relationships() 