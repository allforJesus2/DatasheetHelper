#!/usr/bin/env python3
"""
Main entry point for the Dynamic Datasheet System
"""

import tkinter as tk
import sys
import os

# Add the parent directory to the path so we can import the package
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from dynamic_datasheet_system import DynamicDatasheetApp


def main():
    """Main function to run the application"""
    try:
        # Create the main window
        root = tk.Tk()
        
        # Create the application
        app = DynamicDatasheetApp(root)
        
        # Start the main loop
        root.mainloop()
        
    except Exception as e:
        print(f"Error starting application: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main() 