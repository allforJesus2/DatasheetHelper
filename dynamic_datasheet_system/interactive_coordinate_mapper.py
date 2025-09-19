#!/usr/bin/env python3
"""
Interactive Coordinate Mapper - Allows clicking on Excel cells to map coordinates to data fields
"""

import tkinter as tk
from tkinter import ttk, messagebox
import xlwings as xw
import re
from typing import Dict, Any, Optional, Callable


class InteractiveCoordinateMapper:
    """Interactive coordinate mapping interface"""
    
    def __init__(self, parent, datasheet_path: str, relationship_system, on_save_callback: Optional[Callable] = None):
        self.parent = parent
        self.datasheet_path = datasheet_path
        self.relationship_system = relationship_system
        self.on_save_callback = on_save_callback
        
        # Excel connection
        self.app = None
        self.wb = None
        
        # Coordinate mappings for each data type
        self.coordinate_mappings = {}
        for data_type in self.relationship_system.get_all_data_types():
            self.coordinate_mappings[data_type] = {}
        
        # Top tag (first coordinate for primary data type)
        self.top_tag = None
        
        # Create the window
        self.create_window()
        
    def create_window(self):
        """Create the coordinate mapping window"""
        self.window = tk.Toplevel(self.parent)
        self.window.title("Interactive Coordinate Mapper")
        self.window.geometry("1000x700")
        
        # Make it modal
        self.window.transient(self.parent)
        self.window.grab_set()
        
        # Create widgets
        self.create_widgets()
        
        # Initialize Excel connection
        self.init_excel()
        
        # Start coordinate tracking
        self.window.after(200, self.update_coordinate_entry)
        
    def create_widgets(self):
        """Create the GUI widgets"""
        # Main container
        main_frame = ttk.Frame(self.window)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Title
        title_label = ttk.Label(main_frame, text="Interactive Coordinate Mapping", 
                               font=("Arial", 14, "bold"))
        title_label.pack(pady=(0, 10))
        
        # Instructions
        instructions = ttk.Label(main_frame, 
                               text="1. Open Excel and click on a cell to get its coordinate\n"
                                    "2. Select a data field from the dropdown\n"
                                    "3. Click 'Add' to map the coordinate to that field\n"
                                    "4. Use 'Add/Update Datasheets' to populate the datasheet",
                               font=("Arial", 10))
        instructions.pack(pady=(0, 10))
        
        # Coordinate entry frame
        coord_frame = ttk.LabelFrame(main_frame, text="Coordinate Selection", padding=10)
        coord_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(coord_frame, text="Current Excel Cell:").pack(side=tk.LEFT)
        
        self.coord_var = tk.StringVar()
        self.coord_entry = ttk.Entry(coord_frame, textvariable=self.coord_var, width=15)
        self.coord_entry.pack(side=tk.LEFT, padx=5)
        
        # Increment/decrement buttons
        ttk.Button(coord_frame, text="+", width=3, 
                  command=self.increment_coordinate).pack(side=tk.LEFT, padx=2)
        ttk.Button(coord_frame, text="-", width=3, 
                  command=self.decrement_coordinate).pack(side=tk.LEFT, padx=2)
        
        # Data type mapping frames
        mapping_frame = ttk.Frame(main_frame)
        mapping_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # Create frames for each data type
        self.data_type_frames = {}
        self.combos = {}
        self.listboxes = {}
        
        for data_type in self.relationship_system.get_all_data_types():
            config = self.relationship_system.get_data_type_config(data_type)
            name = config.get('name', data_type.upper())
            short_name = config.get('short_name', data_type.upper())
            
            # Create frame for this data type
            data_frame = ttk.LabelFrame(mapping_frame, text=f"{name} Coordinates")
            data_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)
            
            # Controls frame
            controls = ttk.Frame(data_frame)
            controls.pack(fill=tk.X, padx=5, pady=5)
            
            # Label
            label = ttk.Label(controls, text=f"Select {short_name} Field:")
            label.pack(side=tk.LEFT)
            
            # Combo box with data fields
            combo_values = self.get_combo_values(data_type)
            combo = ttk.Combobox(controls, values=combo_values, state="readonly", width=20)
            combo.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
            
            # Button frame
            btn_frame = ttk.Frame(data_frame)
            btn_frame.pack(fill=tk.X, padx=5, pady=5)
            
            # Listbox to show mappings
            listbox = tk.Listbox(data_frame, height=15)
            listbox.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
            
            # Buttons
            ttk.Button(btn_frame, text=f"Add to {short_name}",
                      command=lambda dt=data_type: self.add_coordinate(dt)).pack(side=tk.LEFT, padx=2)
            ttk.Button(btn_frame, text="Remove",
                      command=lambda dt=data_type: self.remove_coordinate(dt)).pack(side=tk.LEFT, padx=2)
            ttk.Button(btn_frame, text="Clear All",
                      command=lambda dt=data_type: self.clear_coordinates(dt)).pack(side=tk.LEFT, padx=2)
            
            # Store references
            self.data_type_frames[data_type] = data_frame
            self.combos[data_type] = combo
            self.listboxes[data_type] = listbox
            
            # Update listbox
            self.update_listbox(data_type)
        
        # Action buttons frame
        action_frame = ttk.Frame(main_frame)
        action_frame.pack(fill=tk.X, pady=10)
        
        ttk.Button(action_frame, text="Add/Update Datasheets", 
                  command=self.add_update_datasheets).pack(side=tk.LEFT, padx=5)
        ttk.Button(action_frame, text="Save Mappings", 
                  command=self.save_mappings).pack(side=tk.LEFT, padx=5)
        ttk.Button(action_frame, text="Close", 
                  command=self.close_window).pack(side=tk.RIGHT, padx=5)
        
    def get_combo_values(self, data_type: str) -> list:
        """Get combo values for a data type"""
        data = self.relationship_system.get_data(data_type)
        if data:
            # Get the first item's keys
            for key, value in data.items():
                if isinstance(value, dict):
                    return list(value.keys())
        return []
        
    def init_excel(self):
        """Initialize Excel connection"""
        try:
            if self.datasheet_path:
                self.app = xw.App(visible=True)
                self.wb = self.app.books.open(self.datasheet_path)
                messagebox.showinfo("Success", f"Opened Excel file: {self.datasheet_path}")
            else:
                messagebox.showwarning("Warning", "No datasheet path provided")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open Excel file: {e}")
            
    def update_coordinate_entry(self):
        """Update the coordinate entry with current Excel selection"""
        try:
            # Check if Excel app and workbook are still open
            if not self._is_excel_accessible():
                print("Excel application or workbook was closed, closing coordinate mapper...")
                self.close_window()
                return
                
            if self.app and self.app.books.active:
                current_selection = self.app.selection.address
                current_selection = current_selection.split(':')[0].replace('$', '')
                self.coord_var.set(current_selection)
        except Exception as e:
            # Excel might not be active or selection might not be available
            print(f"Error updating coordinate entry: {e}")
            if not self._is_excel_accessible():
                print("Excel appears to be closed, closing coordinate mapper...")
                self.close_window()
                return
        
        # Continue updating
        self.window.after(200, self.update_coordinate_entry)
        
    def increment_coordinate(self):
        """Increment the coordinate (e.g., A1 -> A2)"""
        coord = self.coord_var.get()
        new_coord = self.increment_cell_reference(coord, 1)
        self.coord_var.set(new_coord)
        
    def decrement_coordinate(self):
        """Decrement the coordinate (e.g., A2 -> A1)"""
        coord = self.coord_var.get()
        new_coord = self.increment_cell_reference(coord, -1)
        self.coord_var.set(new_coord)
        
    def increment_cell_reference(self, cell_ref: str, increment: int) -> str:
        """Increment a cell reference by the given amount"""
        # Extract column and row
        match = re.match(r'^([A-Z]+)(\d+)$', cell_ref)
        if match:
            col = match.group(1)
            row = int(match.group(2))
            new_row = row + increment
            if new_row > 0:  # Don't allow row 0 or negative
                return f"{col}{new_row}"
        return cell_ref
        
    def add_coordinate(self, data_type: str):
        """Add a coordinate mapping for a data type"""
        coord = self.coord_var.get()
        field = self.combos[data_type].get()
        
        if coord and field:
            self.coordinate_mappings[data_type][coord] = field
            
            # Set top tag if this is the first coordinate for the primary data type
            primary_type = self.relationship_system.get_primary_data_type()
            if data_type == primary_type and not self.top_tag:
                self.top_tag = coord
                
            self.update_listbox(data_type)
        else:
            messagebox.showwarning("Warning", "Please select both a coordinate and a field")
            
    def remove_coordinate(self, data_type: str):
        """Remove a coordinate mapping for a data type"""
        listbox = self.listboxes[data_type]
        selected = listbox.curselection()
        
        if selected:
            idx = selected[0]
            coord = listbox.get(idx).split(':')[0].strip()
            
            if coord in self.coordinate_mappings[data_type]:
                del self.coordinate_mappings[data_type][coord]
                
                # Update top tag if needed
                if coord == self.top_tag:
                    primary_type = self.relationship_system.get_primary_data_type()
                    if data_type == primary_type:
                        # Set to first available coordinate or None
                        coords = list(self.coordinate_mappings[data_type].keys())
                        self.top_tag = coords[0] if coords else None
                        
                self.update_listbox(data_type)
                
    def clear_coordinates(self, data_type: str):
        """Clear all coordinate mappings for a data type"""
        self.coordinate_mappings[data_type].clear()
        
        # Update top tag if needed
        primary_type = self.relationship_system.get_primary_data_type()
        if data_type == primary_type:
            self.top_tag = None
            
        self.update_listbox(data_type)
        
    def update_listbox(self, data_type: str):
        """Update the listbox for a data type"""
        listbox = self.listboxes[data_type]
        listbox.delete(0, tk.END)
        
        for coord, field in self.coordinate_mappings[data_type].items():
            display_text = f"{coord}: {field}"
            if coord == self.top_tag:
                display_text += " (Top Tag)"
            listbox.insert(tk.END, display_text)
            
    def save_mappings(self):
        """Save coordinate mappings to the relationship system"""
        try:
            # Update coordinate values in the relationship system
            for data_type, mappings in self.coordinate_mappings.items():
                config = self.relationship_system.get_data_type_config(data_type)
                if config:
                    config['coordinate_values'] = mappings.copy()
                    
            # Update top tag
            if self.top_tag:
                primary_type = self.relationship_system.get_primary_data_type()
                if primary_type:
                    config = self.relationship_system.get_data_type_config(primary_type)
                    if config:
                        config['top_tag'] = self.top_tag
                        
            messagebox.showinfo("Success", "Coordinate mappings saved to relationship system")
            
            # Call callback if provided
            if self.on_save_callback:
                self.on_save_callback(self.coordinate_mappings, self.top_tag)
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save mappings: {e}")
            
    def add_update_datasheets(self):
        """Add/Update datasheets using the main_functions.add_update_datasheets"""
        try:
            # Import the function
            import sys
            import os
            sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
            try:
                from main_functions import add_update_datasheets
            except ImportError:
                # Try alternative import path
                sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
                from main_functions import add_update_datasheets
            
            # Get the primary data type's coordinate mappings
            primary_type = self.relationship_system.get_primary_data_type()
            if not primary_type:
                messagebox.showerror("Error", "No primary data type set")
                return
                
            # Get coordinate mappings for the primary data type
            primary_mappings = self.coordinate_mappings.get(primary_type, {})
            if not primary_mappings:
                messagebox.showerror("Error", f"No coordinate mappings for {primary_type}")
                return
                
            # Get data for the primary data type
            primary_data = self.relationship_system.get_data(primary_type)
            if not primary_data:
                messagebox.showerror("Error", f"No data available for {primary_type}")
                return
                
            # Generate tag_cell_values from the data and mappings
            tag_cell_values = {}
            for tag_key, tag_data in primary_data.items():
                cell_values = {}
                for coord, field in primary_mappings.items():
                    if field in tag_data:
                        cell_values[coord] = tag_data[field]
                if cell_values:
                    tag_cell_values[tag_key] = cell_values
                    
            if not tag_cell_values:
                messagebox.showerror("Error", "No valid tag-cell values generated")
                return
                
            # Call add_update_datasheets
            # Note: You'll need to provide the required parameters
            # This is a simplified call - you may need to adjust based on your needs
            result = add_update_datasheets(
                datasheet=self.wb,
                source_sheet_name=None,  # You can make this configurable
                tag_cell_values=tag_cell_values,
                datasheet_coord="A1",  # You can make this configurable
                ds_prefix="DS",  # You can make this configurable
                rows_per_sheet=1,
                key_coordinate=self.top_tag or "A1"
            )
            
            messagebox.showinfo("Success", "Datasheets updated successfully")
            
        except ImportError:
            messagebox.showerror("Error", "Could not import add_update_datasheets from main_functions")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to update datasheets: {e}")
            
    def _is_excel_accessible(self):
        """Check if Excel application and workbook are still accessible"""
        try:
            if not self.app:
                return False
            
            # Try to access the app's books collection
            # This will raise an exception if Excel is closed
            _ = self.app.books
            return True
        except Exception:
            # Excel application is closed or no longer accessible
            return False

    def close_window(self):
        """Close the window"""
        try:
            if self.wb:
                self.wb.save()
            if self.app:
                self.app.quit()
        except:
            pass
        self.window.destroy() 