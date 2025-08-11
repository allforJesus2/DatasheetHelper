#!/usr/bin/env python3
"""
Dynamic Datasheet Application - Main GUI application for the dynamic datasheet system
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import os
import json
from typing import Dict, Any, Optional, List

try:
    from .dynamic_relationship_system import DynamicRelationshipSystem
    from .excel_manager import ExcelManager
except ImportError:
    from dynamic_relationship_system import DynamicRelationshipSystem
    from excel_manager import ExcelManager


class DynamicDatasheetApp:
    """Main application class for the dynamic datasheet system"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("Dynamic Datasheet System")
        self.root.geometry("1200x800")
        
        # Initialize core systems
        self.relationship_system = DynamicRelationshipSystem()
        self.excel_manager = ExcelManager()
        
        # GUI state
        self.file_paths = {}  # Maps data_type -> filepath
        self.destination_path = ""
        
        # Initialize with default data types
        self._initialize_default_data_types()
        
        # Create GUI
        self.create_widgets()
        
    def _initialize_default_data_types(self):
        """Initialize with some default data types"""
        # Tag Dictionary (TD)
        td_config = {
            'name': 'Tag Dictionary',
            'short_name': 'TD',
            'description': 'Tag data from Excel files',
            'default_headers': ['TAG NUMBER', 'DESCRIPTION', 'LOCATION'],
            'coordinate_values': {'A1': 'TAG NUMBER', 'B1': 'DESCRIPTION'},
            'selected_sheets': None,
            'path_key': 'td_path'
        }
        self.relationship_system.add_data_type('td', td_config)
        
        # Process Conditions (PC)
        pc_config = {
            'name': 'Process Conditions',
            'short_name': 'PC', 
            'description': 'Process conditions from Excel files',
            'default_headers': ['Line No.', 'Pressure', 'Temperature'],
            'coordinate_values': {'C1': 'Line No.', 'D1': 'Pressure'},
            'selected_sheets': None,
            'path_key': 'pc_path'
        }
        self.relationship_system.add_data_type('pc', pc_config)
        
        # Set TD as primary (can be changed)
        self.relationship_system.set_primary_data_type('td')
        
        # Add a sample relationship
        self.relationship_system.add_relationship('td', 'pc', {
            'source_key': 'TAG NUMBER',
            'target_key': 'Line No.',
            'transformation': 'x.split("-")[2] if "-" in x else x',
            'description': 'TD references PC via Line Number'
        })
        
    def create_widgets(self):
        """Create the main GUI widgets"""
        # Create main notebook
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Main tab
        self.create_main_tab()
        
        # Relationships tab
        self.create_relationships_tab()
        
        # Data tab
        self.create_data_tab()
        
        # Coordinates tab
        self.create_coordinates_tab()
        
        # Settings tab
        self.create_settings_tab()
        
    def create_main_tab(self):
        """Create the main tab with data sources and destination"""
        main_frame = ttk.Frame(self.notebook)
        self.notebook.add(main_frame, text="Main")
        
        # Title
        title_label = ttk.Label(main_frame, text="Dynamic Datasheet System", 
                               font=("Arial", 16, "bold"))
        title_label.pack(pady=10)
        
        # Data Sources Section
        sources_frame = ttk.LabelFrame(main_frame, text="Data Sources", padding=10)
        sources_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # Sources label
        sources_label = ttk.Label(sources_frame, text="DATA SOURCES", 
                                 font=("Arial", 12, "bold"), foreground="green")
        sources_label.pack(anchor=tk.W, pady=(0, 10))
        
        # Create data type frames
        self.create_data_type_frames(sources_frame)
        
        # Separator
        separator = ttk.Separator(main_frame, orient='horizontal')
        separator.pack(fill=tk.X, padx=10, pady=10)
        
        # Destination Section
        destination_frame = ttk.LabelFrame(main_frame, text="Destination", padding=10)
        destination_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # Destination label
        dest_label = ttk.Label(destination_frame, text="DESTINATION", 
                              font=("Arial", 12, "bold"), foreground="blue")
        dest_label.pack(anchor=tk.W, pady=(0, 10))
        
        # Destination file selection
        dest_container = ttk.Frame(destination_frame)
        dest_container.pack(fill=tk.X)
        
        ttk.Label(dest_container, text="Datasheet Output:").pack(side=tk.LEFT)
        
        self.dest_entry = ttk.Entry(dest_container, width=50)
        self.dest_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        ttk.Button(dest_container, text="Browse", 
                  command=self.browse_destination).pack(side=tk.LEFT, padx=5)
        
        # Action buttons
        action_frame = ttk.Frame(main_frame)
        action_frame.pack(pady=20)
        
        ttk.Button(action_frame, text="Process Coordinates", 
                  command=self.process_coordinates).pack(side=tk.LEFT, padx=5)
        ttk.Button(action_frame, text="Save Configuration", 
                  command=self.save_configuration).pack(side=tk.LEFT, padx=5)
        ttk.Button(action_frame, text="Load Configuration", 
                  command=self.load_configuration).pack(side=tk.LEFT, padx=5)
        
    def create_data_type_frames(self, parent):
        """Create frames for each data type"""
        self.data_type_frames = {}
        
        for data_type in self.relationship_system.get_all_data_types():
            self.create_single_data_type_frame(parent, data_type)
            
    def create_single_data_type_frame(self, parent, data_type):
        """Create a frame for a single data type"""
        config = self.relationship_system.get_data_type_config(data_type)
        name = config.get('name', data_type.upper())
        
        # Create frame
        frame = ttk.LabelFrame(parent, text=name, padding=5)
        frame.pack(fill=tk.X, pady=5)
        
        # File path
        path_frame = ttk.Frame(frame)
        path_frame.pack(fill=tk.X, pady=2)
        
        ttk.Label(path_frame, text="File:").pack(side=tk.LEFT)
        
        entry = ttk.Entry(path_frame, width=40)
        entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        ttk.Button(path_frame, text="Browse", 
                  command=lambda: self.browse_file(data_type, entry)).pack(side=tk.LEFT, padx=2)
        
        # Action buttons
        button_frame = ttk.Frame(frame)
        button_frame.pack(fill=tk.X, pady=2)
        
        ttk.Button(button_frame, text="Load Data", 
                  command=lambda: self.load_data_type(data_type)).pack(side=tk.LEFT, padx=2)
        ttk.Button(button_frame, text="Configure", 
                  command=lambda: self.configure_data_type(data_type)).pack(side=tk.LEFT, padx=2)
        ttk.Button(button_frame, text="View Data", 
                  command=lambda: self.view_data_type(data_type)).pack(side=tk.LEFT, padx=2)
        
        # Store references
        self.data_type_frames[data_type] = {
            'frame': frame,
            'entry': entry
        }
        
    def create_relationships_tab(self):
        """Create the relationships configuration tab"""
        relationships_frame = ttk.Frame(self.notebook)
        self.notebook.add(relationships_frame, text="Relationships")
        
        # Title
        title_label = ttk.Label(relationships_frame, text="Data Type Relationships", 
                               font=("Arial", 14, "bold"))
        title_label.pack(pady=10)
        
        # Primary data type selection
        primary_frame = ttk.LabelFrame(relationships_frame, text="Primary Data Type", padding=10)
        primary_frame.pack(fill=tk.X, padx=10, pady=5)
        
        ttk.Label(primary_frame, text="Select Primary Data Type:").pack(side=tk.LEFT)
        
        self.primary_var = tk.StringVar()
        self.primary_combo = ttk.Combobox(primary_frame, textvariable=self.primary_var, 
                                    state="readonly")
        self.primary_combo.pack(side=tk.LEFT, padx=5)
        self.primary_combo.bind('<<ComboboxSelected>>', self.on_primary_changed)
        
        # Update primary combo
        self.update_primary_combo()
        
        # Relationships list
        relationships_list_frame = ttk.LabelFrame(relationships_frame, text="Current Relationships", padding=10)
        relationships_list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # Create treeview for relationships
        columns = ('Source', 'Target', 'Source Key', 'Target Key', 'Transformation')
        self.relationships_tree = ttk.Treeview(relationships_list_frame, columns=columns, show='headings')
        
        for col in columns:
            self.relationships_tree.heading(col, text=col)
            self.relationships_tree.column(col, width=120)
            
        self.relationships_tree.pack(fill=tk.BOTH, expand=True)
        
        # Add relationship button
        ttk.Button(relationships_list_frame, text="Add Relationship", 
                  command=self.add_relationship_dialog).pack(pady=5)
        
        # Update relationships display
        self.update_relationships_display()
        
    def create_data_tab(self):
        """Create the data management tab"""
        data_frame = ttk.Frame(self.notebook)
        self.notebook.add(data_frame, text="Data")
        
        # Title
        title_label = ttk.Label(data_frame, text="Data Management", 
                               font=("Arial", 14, "bold"))
        title_label.pack(pady=10)
        
        # Data type selector
        selector_frame = ttk.Frame(data_frame)
        selector_frame.pack(fill=tk.X, padx=10, pady=5)
        
        ttk.Label(selector_frame, text="Select Data Type:").pack(side=tk.LEFT)
        
        self.data_type_var = tk.StringVar()
        self.data_type_combo = ttk.Combobox(selector_frame, textvariable=self.data_type_var, 
                                           state="readonly")
        self.data_type_combo.pack(side=tk.LEFT, padx=5)
        self.data_type_combo.bind('<<ComboboxSelected>>', self.on_data_type_selected)
        
        # Update data type combo
        self.update_data_type_combo()
        
        # Data display
        data_display_frame = ttk.LabelFrame(data_frame, text="Data", padding=10)
        data_display_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        self.data_text = scrolledtext.ScrolledText(data_display_frame, height=20)
        self.data_text.pack(fill=tk.BOTH, expand=True)
        
    def create_coordinates_tab(self):
        """Create the coordinates processing tab"""
        coordinates_frame = ttk.Frame(self.notebook)
        self.notebook.add(coordinates_frame, text="Coordinates")
        
        # Title
        title_label = ttk.Label(coordinates_frame, text="Coordinate Processing", 
                               font=("Arial", 14, "bold"))
        title_label.pack(pady=10)
        
        # Instructions
        instructions = ttk.Label(coordinates_frame, 
                               text="Use the Interactive Coordinate Mapper to map Excel cells to data fields,\n"
                                    "then process coordinates and update datasheets.",
                               font=("Arial", 10))
        instructions.pack(pady=(0, 10))
        
        # Button frame
        button_frame = ttk.Frame(coordinates_frame)
        button_frame.pack(pady=10)
        
        # Interactive mapping button
        ttk.Button(button_frame, text="Interactive Coordinate Mapper", 
                  command=self.open_interactive_mapper).pack(side=tk.LEFT, padx=5)
        
        # Process button
        ttk.Button(button_frame, text="Process Coordinates", 
                  command=self.process_coordinates).pack(side=tk.LEFT, padx=5)
        
        # Results display
        results_frame = ttk.LabelFrame(coordinates_frame, text="Results", padding=10)
        results_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        self.coordinates_text = scrolledtext.ScrolledText(results_frame, height=20)
        self.coordinates_text.pack(fill=tk.BOTH, expand=True)
        
    def create_settings_tab(self):
        """Create the settings tab"""
        settings_frame = ttk.Frame(self.notebook)
        self.notebook.add(settings_frame, text="Settings")
        
        # Title
        title_label = ttk.Label(settings_frame, text="System Settings", 
                               font=("Arial", 14, "bold"))
        title_label.pack(pady=10)
        
        # Existing data types section
        existing_frame = ttk.LabelFrame(settings_frame, text="Existing Data Types", padding=10)
        existing_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # Create treeview for data types
        columns = ('Key', 'Name', 'Short Name', 'Description', 'Status')
        self.data_types_tree = ttk.Treeview(existing_frame, columns=columns, show='headings', height=8)
        
        # Configure columns
        for col in columns:
            self.data_types_tree.heading(col, text=col)
            self.data_types_tree.column(col, width=120, minwidth=100)
        
        # Add scrollbar
        tree_scrollbar = ttk.Scrollbar(existing_frame, orient=tk.VERTICAL, command=self.data_types_tree.yview)
        self.data_types_tree.configure(yscrollcommand=tree_scrollbar.set)
        
        # Pack treeview and scrollbar
        self.data_types_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        tree_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Buttons frame
        buttons_frame = ttk.Frame(existing_frame)
        buttons_frame.pack(fill=tk.X, pady=(10, 0))
        
        ttk.Button(buttons_frame, text="Configure Selected", 
                  command=self.configure_selected_data_type).pack(side=tk.LEFT, padx=5)
        ttk.Button(buttons_frame, text="Remove Selected", 
                  command=self.remove_selected_data_type).pack(side=tk.LEFT, padx=5)
        ttk.Button(buttons_frame, text="Refresh List", 
                  command=self.refresh_data_types_list).pack(side=tk.LEFT, padx=5)
        
        # Add new data type
        add_frame = ttk.LabelFrame(settings_frame, text="Add New Data Type", padding=10)
        add_frame.pack(fill=tk.X, padx=10, pady=5)
        
        ttk.Button(add_frame, text="Add Data Type", 
                  command=self.add_data_type_dialog).pack(pady=5)
        
        # System validation
        validation_frame = ttk.LabelFrame(settings_frame, text="System Validation", padding=10)
        validation_frame.pack(fill=tk.X, padx=10, pady=5)
        
        ttk.Button(validation_frame, text="Validate System", 
                  command=self.validate_system).pack(pady=5)
        
        # Validation results
        self.validation_text = scrolledtext.ScrolledText(validation_frame, height=8)
        self.validation_text.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # Initialize the data types list
        self.refresh_data_types_list()
        
    # Event handlers and utility methods
    def browse_file(self, data_type, entry):
        """Browse for a file for a specific data type"""
        filename = filedialog.askopenfilename(
            title=f"Select file for {data_type}",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if filename:
            entry.delete(0, tk.END)
            entry.insert(0, filename)
            self.file_paths[data_type] = filename
            
    def browse_destination(self):
        """Browse for destination file"""
        filename = filedialog.asksaveasfilename(
            title="Select destination file",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if filename:
            self.dest_entry.delete(0, tk.END)
            self.dest_entry.insert(0, filename)
            self.destination_path = filename
            
    def load_data_type(self, data_type):
        """Load data for a specific data type"""
        if data_type not in self.file_paths:
            messagebox.showerror("Error", f"No file selected for {data_type}")
            return
            
        filepath = self.file_paths[data_type]
        config = self.relationship_system.get_data_type_config(data_type)
        
        try:
            # Load data from Excel
            data = self.excel_manager.generate_dictionary_from_xlsx(
                filepath, 
                sheet_names=config.get('selected_sheets'),
                key_column=config.get('default_headers', [None])[0] if config.get('default_headers') else None
            )
            
            # Set data in relationship system
            self.relationship_system.set_data(data_type, data)
            
            messagebox.showinfo("Success", f"Loaded {len(data)} items for {data_type}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load data for {data_type}: {e}")
            
    def configure_data_type(self, data_type):
        """Configure a data type (headers, sheets, etc.)"""
        self.open_data_type_config_dialog(data_type)
        
    def open_data_type_config_dialog(self, data_type):
        """Open configuration dialog for a data type"""
        config = self.relationship_system.get_data_type_config(data_type)
        name = config.get('name', data_type.upper())
        
        # Create dialog
        dialog = tk.Toplevel(self.root)
        dialog.title(f"Configure {name}")
        dialog.geometry("600x500")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Main frame
        main_frame = ttk.Frame(dialog, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Title
        title_label = ttk.Label(main_frame, text=f"Configure {name}", font=("Arial", 14, "bold"))
        title_label.pack(pady=(0, 10))
        
        # Configuration fields
        fields_frame = ttk.LabelFrame(main_frame, text="Configuration", padding="10")
        fields_frame.pack(fill=tk.X, pady=5)
        
        # Name field
        ttk.Label(fields_frame, text="Display Name:").grid(row=0, column=0, sticky=tk.W, pady=2)
        name_var = tk.StringVar(value=config.get('name', data_type.upper()))
        name_entry = ttk.Entry(fields_frame, textvariable=name_var, width=30)
        name_entry.grid(row=0, column=1, sticky=tk.EW, padx=5, pady=2)
        
        # Short name field
        ttk.Label(fields_frame, text="Short Name:").grid(row=1, column=0, sticky=tk.W, pady=2)
        short_name_var = tk.StringVar(value=config.get('short_name', data_type.upper()))
        short_name_entry = ttk.Entry(fields_frame, textvariable=short_name_var, width=30)
        short_name_entry.grid(row=1, column=1, sticky=tk.EW, padx=5, pady=2)
        
        # File path field
        ttk.Label(fields_frame, text="File Path:").grid(row=2, column=0, sticky=tk.W, pady=2)
        path_var = tk.StringVar(value=config.get('path', ''))
        path_entry = ttk.Entry(fields_frame, textvariable=path_var, width=30)
        path_entry.grid(row=2, column=1, sticky=tk.EW, padx=5, pady=2)
        
        # Browse button
        ttk.Button(fields_frame, text="Browse", 
                  command=lambda: self.browse_file_for_config(path_var)).grid(row=2, column=2, padx=5, pady=2)
        
        # Selected sheets field
        ttk.Label(fields_frame, text="Selected Sheets:").grid(row=3, column=0, sticky=tk.W, pady=2)
        sheets_var = tk.StringVar(value=', '.join(config.get('selected_sheets') or []))
        sheets_entry = ttk.Entry(fields_frame, textvariable=sheets_var, width=30)
        sheets_entry.grid(row=3, column=1, sticky=tk.EW, padx=5, pady=2)
        
        # Headers field
        ttk.Label(fields_frame, text="Headers:").grid(row=4, column=0, sticky=tk.W, pady=2)
        headers_var = tk.StringVar(value=', '.join(config.get('headers') or []))
        headers_entry = ttk.Entry(fields_frame, textvariable=headers_var, width=30)
        headers_entry.grid(row=4, column=1, sticky=tk.EW, padx=5, pady=2)
        
        # Configure grid weights
        fields_frame.columnconfigure(1, weight=1)
        
        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=10)
        
        def save_config():
            try:
                # Update configuration
                new_config = config.copy()
                new_config['name'] = name_var.get()
                new_config['short_name'] = short_name_var.get()
                new_config['path'] = path_var.get()
                new_config['selected_sheets'] = [s.strip() for s in sheets_var.get().split(',') if s.strip()]
                new_config['headers'] = [h.strip() for h in headers_var.get().split(',') if h.strip()]
                
                # Update in relationship system
                self.relationship_system.update_data_type_config(data_type, new_config)
                
                messagebox.showinfo("Success", f"Configuration saved for {name}")
                dialog.destroy()
                
                # Refresh GUI
                self.refresh_gui()
                
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save configuration: {e}")
        
        ttk.Button(button_frame, text="Save", command=save_config).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Cancel", command=dialog.destroy).pack(side=tk.RIGHT, padx=5)
        
    def browse_file_for_config(self, path_var):
        """Browse for file in configuration dialog"""
        filename = filedialog.askopenfilename(
            title="Select File",
            filetypes=[("Excel files", "*.xlsx *.xlsm"), ("All files", "*.*")]
        )
        if filename:
            path_var.set(filename)
        
    def view_data_type(self, data_type):
        """View data for a specific data type"""
        data = self.relationship_system.get_data(data_type)
        config = self.relationship_system.get_data_type_config(data_type)
        name = config.get('name', data_type.upper())
        
        # Create a new window to display data
        view_window = tk.Toplevel(self.root)
        view_window.title(f"Data: {name}")
        view_window.geometry("800x600")
        
        # Display data
        text_widget = scrolledtext.ScrolledText(view_window)
        text_widget.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        text_widget.insert(tk.END, f"Data for {name}:\n\n")
        text_widget.insert(tk.END, json.dumps(data, indent=2))
        text_widget.config(state=tk.DISABLED)
        
    def open_interactive_mapper(self):
        """Open the interactive coordinate mapper"""
        if not self.destination_path:
            messagebox.showerror("Error", "Please select a destination datasheet first")
            return
            
        try:
            # Import the interactive mapper
            from .interactive_coordinate_mapper import InteractiveCoordinateMapper
            
            # Create the mapper
            mapper = InteractiveCoordinateMapper(
                parent=self.root,
                datasheet_path=self.destination_path,
                relationship_system=self.relationship_system,
                on_save_callback=self.on_coordinate_mappings_saved
            )
            
        except ImportError:
            messagebox.showerror("Error", "Could not import InteractiveCoordinateMapper")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open coordinate mapper: {e}")
            
    def on_coordinate_mappings_saved(self, coordinate_mappings, top_tag):
        """Callback when coordinate mappings are saved"""
        print(f"Coordinate mappings saved: {coordinate_mappings}")
        print(f"Top tag: {top_tag}")
        
        # Update the relationship system with the new mappings
        for data_type, mappings in coordinate_mappings.items():
            config = self.relationship_system.get_data_type_config(data_type)
            if config:
                config['coordinate_values'] = mappings.copy()
                
        messagebox.showinfo("Success", "Coordinate mappings updated in the system")
        
    def process_coordinates(self):
        """Process coordinates using the relationship system"""
        try:
            # Process coordinates
            results = self.relationship_system.process_coordinates_dynamic()
            
            # Display results
            self.coordinates_text.delete(1.0, tk.END)
            self.coordinates_text.insert(tk.END, "Coordinate Processing Results:\n\n")
            self.coordinates_text.insert(tk.END, json.dumps(results, indent=2))
            
            # Save to destination if specified
            if self.destination_path and results:
                self.excel_manager.write_coordinate_data_to_excel(
                    results, self.destination_path
                )
                messagebox.showinfo("Success", f"Results saved to {self.destination_path}")
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to process coordinates: {e}")
            
    def save_configuration(self):
        """Save the current configuration"""
        filename = filedialog.asksaveasfilename(
            title="Save Configuration",
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        if filename:
            if self.relationship_system.save_configuration(filename):
                messagebox.showinfo("Success", f"Configuration saved to {filename}")
            else:
                messagebox.showerror("Error", "Failed to save configuration")
                
    def load_configuration(self):
        """Load a configuration"""
        filename = filedialog.askopenfilename(
            title="Load Configuration",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        if filename:
            if self.relationship_system.load_configuration(filename):
                messagebox.showinfo("Success", f"Configuration loaded from {filename}")
                # Refresh GUI
                self.refresh_gui()
            else:
                messagebox.showerror("Error", "Failed to load configuration")
                
    def update_primary_combo(self):
        """Update the primary data type combo box"""
        data_types = self.relationship_system.get_all_data_types()
        names = self.relationship_system.get_data_type_names()
        
        self.primary_combo['values'] = [names.get(dt, dt.upper()) for dt in data_types]
        
        # Set current primary
        current_primary = self.relationship_system.get_primary_data_type()
        if current_primary:
            self.primary_var.set(names.get(current_primary, current_primary.upper()))
            
    def on_primary_changed(self, event):
        """Handle primary data type change"""
        selected_name = self.primary_var.get()
        names = self.relationship_system.get_data_type_names()
        
        # Find data type key for selected name
        for dt, name in names.items():
            if name == selected_name:
                self.relationship_system.set_primary_data_type(dt)
                break
                
    def update_relationships_display(self):
        """Update the relationships treeview"""
        # Clear existing items
        for item in self.relationships_tree.get_children():
            self.relationships_tree.delete(item)
            
        # Add current relationships
        relationships = self.relationship_system.get_relationships()
        names = self.relationship_system.get_data_type_names()
        
        for source_type, targets in relationships.items():
            for target_type, config in targets.items():
                self.relationships_tree.insert('', 'end', values=(
                    names.get(source_type, source_type.upper()),
                    names.get(target_type, target_type.upper()),
                    config.get('source_key', ''),
                    config.get('target_key', ''),
                    config.get('transformation', '')
                ))
                
    def add_relationship_dialog(self):
        """Open dialog to add a new relationship"""
        self.open_relationship_config_dialog()
        
    def open_relationship_config_dialog(self):
        """Open dialog to configure relationships"""
        # Create dialog
        dialog = tk.Toplevel(self.root)
        dialog.title("Add Relationship")
        dialog.geometry("700x600")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Main frame
        main_frame = ttk.Frame(dialog, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Title
        title_label = ttk.Label(main_frame, text="Add Relationship", font=("Arial", 14, "bold"))
        title_label.pack(pady=(0, 10))
        
        # Configuration frame
        config_frame = ttk.LabelFrame(main_frame, text="Relationship Configuration", padding="10")
        config_frame.pack(fill=tk.X, pady=5)
        
        # Get available data types
        data_types = self.relationship_system.get_all_data_types()
        names = self.relationship_system.get_data_type_names()
        data_type_names = [names.get(dt, dt.upper()) for dt in data_types]
        
        # Source data type
        ttk.Label(config_frame, text="Source Data Type:").grid(row=0, column=0, sticky=tk.W, pady=2)
        source_var = tk.StringVar()
        source_combo = ttk.Combobox(config_frame, textvariable=source_var, values=data_type_names, state="readonly", width=25)
        source_combo.grid(row=0, column=1, sticky=tk.EW, padx=5, pady=2)
        
        # Target data type
        ttk.Label(config_frame, text="Target Data Type:").grid(row=1, column=0, sticky=tk.W, pady=2)
        target_var = tk.StringVar()
        target_combo = ttk.Combobox(config_frame, textvariable=target_var, values=data_type_names, state="readonly", width=25)
        target_combo.grid(row=1, column=1, sticky=tk.EW, padx=5, pady=2)
        
        # Source key
        ttk.Label(config_frame, text="Source Key:").grid(row=2, column=0, sticky=tk.W, pady=2)
        source_key_var = tk.StringVar()
        source_key_entry = ttk.Entry(config_frame, textvariable=source_key_var, width=25)
        source_key_entry.grid(row=2, column=1, sticky=tk.EW, padx=5, pady=2)
        
        # Target key
        ttk.Label(config_frame, text="Target Key:").grid(row=3, column=0, sticky=tk.W, pady=2)
        target_key_var = tk.StringVar()
        target_key_entry = ttk.Entry(config_frame, textvariable=target_key_var, width=25)
        target_key_entry.grid(row=3, column=1, sticky=tk.EW, padx=5, pady=2)
        
        # Transformation
        ttk.Label(config_frame, text="Transformation (optional):").grid(row=4, column=0, sticky=tk.W, pady=2)
        transformation_var = tk.StringVar()
        transformation_entry = ttk.Entry(config_frame, textvariable=transformation_var, width=25)
        transformation_entry.grid(row=4, column=1, sticky=tk.EW, padx=5, pady=2)
        
        # Help text for transformation
        help_text = ttk.Label(config_frame, text="Use 'value' to reference the source value. Example: value * 1.8 + 32", 
                             font=("Arial", 8), foreground="gray")
        help_text.grid(row=5, column=0, columnspan=2, sticky=tk.W, pady=2)
        
        # Configure grid weights
        config_frame.columnconfigure(1, weight=1)
        
        # Available keys display
        keys_frame = ttk.LabelFrame(main_frame, text="Available Keys", padding="10")
        keys_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # Create notebook for different data types
        keys_notebook = ttk.Notebook(keys_frame)
        keys_notebook.pack(fill=tk.BOTH, expand=True)
        
        # Create tabs for each data type
        keys_tabs = {}
        for data_type in data_types:
            name = names.get(data_type, data_type.upper())
            tab = ttk.Frame(keys_notebook)
            keys_notebook.add(tab, text=name)
            
            # Get data for this type
            data = self.relationship_system.get_data(data_type)
            if data:
                # Get keys from first item
                for key, value in data.items():
                    if isinstance(value, dict):
                        keys = list(value.keys())
                        break
                else:
                    keys = []
            else:
                keys = []
            
            # Display keys
            keys_text = scrolledtext.ScrolledText(tab, height=8)
            keys_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
            keys_text.insert(tk.END, f"Available keys for {name}:\n\n")
            for key in keys:
                keys_text.insert(tk.END, f"• {key}\n")
            keys_text.config(state=tk.DISABLED)
            
            keys_tabs[data_type] = tab
        
        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=10)
        
        def save_relationship():
            try:
                # Get selected data types
                source_name = source_var.get()
                target_name = target_var.get()
                
                if not source_name or not target_name:
                    messagebox.showerror("Error", "Please select both source and target data types")
                    return
                
                # Find data type keys
                source_type = None
                target_type = None
                for dt, name in names.items():
                    if name == source_name:
                        source_type = dt
                    if name == target_name:
                        target_type = dt
                
                if not source_type or not target_type:
                    messagebox.showerror("Error", "Could not find data type keys")
                    return
                
                # Validate keys
                source_key = source_key_var.get().strip()
                target_key = target_key_var.get().strip()
                
                if not source_key or not target_key:
                    messagebox.showerror("Error", "Please provide both source and target keys")
                    return
                
                # Create relationship configuration
                relationship_config = {
                    'source_key': source_key,
                    'target_key': target_key
                }
                
                # Add transformation if provided
                transformation = transformation_var.get().strip()
                if transformation:
                    relationship_config['transformation'] = transformation
                
                # Add relationship
                self.relationship_system.add_relationship(source_type, target_type, relationship_config)
                
                messagebox.showinfo("Success", f"Relationship added: {source_name} -> {target_name}")
                dialog.destroy()
                
                # Refresh relationships display
                self.update_relationships_display()
                
            except Exception as e:
                messagebox.showerror("Error", f"Failed to add relationship: {e}")
        
        ttk.Button(button_frame, text="Add Relationship", command=save_relationship).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Cancel", command=dialog.destroy).pack(side=tk.RIGHT, padx=5)
        
    def update_data_type_combo(self):
        """Update the data type combo box"""
        data_types = self.relationship_system.get_all_data_types()
        names = self.relationship_system.get_data_type_names()
        
        self.data_type_combo['values'] = [names.get(dt, dt.upper()) for dt in data_types]
        
    def on_data_type_selected(self, event):
        """Handle data type selection"""
        selected_name = self.data_type_var.get()
        names = self.relationship_system.get_data_type_names()
        
        # Find data type key for selected name
        for dt, name in names.items():
            if name == selected_name:
                # Display data for this type
                data = self.relationship_system.get_data(dt)
                self.data_text.delete(1.0, tk.END)
                self.data_text.insert(tk.END, f"Data for {name}:\n\n")
                self.data_text.insert(tk.END, json.dumps(data, indent=2))
                break
                
    def add_data_type_dialog(self):
        """Open dialog to add a new data type"""
        self.open_add_data_type_dialog()
        
    def open_add_data_type_dialog(self):
        """Open dialog to add a new data type"""
        # Create dialog
        dialog = tk.Toplevel(self.root)
        dialog.title("Add Data Type")
        dialog.geometry("500x400")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Main frame
        main_frame = ttk.Frame(dialog, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Title
        title_label = ttk.Label(main_frame, text="Add New Data Type", font=("Arial", 14, "bold"))
        title_label.pack(pady=(0, 10))
        
        # Configuration frame
        config_frame = ttk.LabelFrame(main_frame, text="Data Type Configuration", padding="10")
        config_frame.pack(fill=tk.X, pady=5)
        
        # Data type key (internal identifier)
        ttk.Label(config_frame, text="Data Type Key:").grid(row=0, column=0, sticky=tk.W, pady=2)
        key_var = tk.StringVar()
        key_entry = ttk.Entry(config_frame, textvariable=key_var, width=30)
        key_entry.grid(row=0, column=1, sticky=tk.EW, padx=5, pady=2)
        
        # Help text for key
        key_help = ttk.Label(config_frame, text="Internal identifier (e.g., 'instrument_index', 'equipment_list')", 
                            font=("Arial", 8), foreground="gray")
        key_help.grid(row=1, column=0, columnspan=2, sticky=tk.W, pady=2)
        
        # Display name
        ttk.Label(config_frame, text="Display Name:").grid(row=2, column=0, sticky=tk.W, pady=2)
        name_var = tk.StringVar()
        name_entry = ttk.Entry(config_frame, textvariable=name_var, width=30)
        name_entry.grid(row=2, column=1, sticky=tk.EW, padx=5, pady=2)
        
        # Short name
        ttk.Label(config_frame, text="Short Name:").grid(row=3, column=0, sticky=tk.W, pady=2)
        short_name_var = tk.StringVar()
        short_name_entry = ttk.Entry(config_frame, textvariable=short_name_var, width=30)
        short_name_entry.grid(row=3, column=1, sticky=tk.EW, padx=5, pady=2)
        
        # File path
        ttk.Label(config_frame, text="File Path (optional):").grid(row=4, column=0, sticky=tk.W, pady=2)
        path_var = tk.StringVar()
        path_entry = ttk.Entry(config_frame, textvariable=path_var, width=30)
        path_entry.grid(row=4, column=1, sticky=tk.EW, padx=5, pady=2)
        
        # Browse button
        ttk.Button(config_frame, text="Browse", 
                  command=lambda: self.browse_file_for_config(path_var)).grid(row=4, column=2, padx=5, pady=2)
        
        # Configure grid weights
        config_frame.columnconfigure(1, weight=1)
        
        # Example data types
        example_frame = ttk.LabelFrame(main_frame, text="Example Data Types", padding="10")
        example_frame.pack(fill=tk.X, pady=5)
        
        examples = [
            ("td", "Tag Dictionary", "TD"),
            ("pc", "Process Conditions", "PC"),
            ("instrument_index", "Instrument Index", "II"),
            ("equipment_list", "Equipment List", "EL"),
            ("piping_specs", "Piping Specifications", "PS")
        ]
        
        for i, (key, name, short) in enumerate(examples):
            example_text = f"{key} → {name} ({short})"
            ttk.Label(example_frame, text=example_text, font=("Arial", 9)).pack(anchor=tk.W, pady=1)
        
        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=10)
        
        def save_data_type():
            try:
                # Validate inputs
                key = key_var.get().strip().lower()
                name = name_var.get().strip()
                short_name = short_name_var.get().strip()
                path = path_var.get().strip()
                
                if not key:
                    messagebox.showerror("Error", "Please provide a data type key")
                    return
                
                if not name:
                    messagebox.showerror("Error", "Please provide a display name")
                    return
                
                if not short_name:
                    messagebox.showerror("Error", "Please provide a short name")
                    return
                
                # Check if data type already exists
                existing_types = self.relationship_system.get_all_data_types()
                if key in existing_types:
                    messagebox.showerror("Error", f"Data type '{key}' already exists")
                    return
                
                # Create configuration
                config = {
                    'name': name,
                    'short_name': short_name,
                    'path': path,
                    'selected_sheets': [],
                    'headers': [],
                    'coordinate_values': {},
                    'data': {}
                }
                
                # Add data type
                self.relationship_system.add_data_type(key, config)
                
                messagebox.showinfo("Success", f"Data type '{name}' added successfully")
                dialog.destroy()
                
                # Refresh GUI
                self.refresh_gui()
                
            except Exception as e:
                messagebox.showerror("Error", f"Failed to add data type: {e}")
        
        ttk.Button(button_frame, text="Add Data Type", command=save_data_type).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Cancel", command=dialog.destroy).pack(side=tk.RIGHT, padx=5)
        
    def validate_system(self):
        """Validate the system configuration"""
        issues = self.relationship_system.validate_system()
        
        self.validation_text.delete(1.0, tk.END)
        
        if not issues:
            self.validation_text.insert(tk.END, "System validation passed!\n")
            self.validation_text.insert(tk.END, "No issues found.\n")
        else:
            self.validation_text.insert(tk.END, "System validation failed!\n\n")
            self.validation_text.insert(tk.END, "Issues found:\n")
            for issue in issues:
                self.validation_text.insert(tk.END, f"- {issue}\n")
                
    def refresh_gui(self):
        """Refresh the GUI after configuration changes"""
        # Update all combo boxes and displays
        self.update_primary_combo()
        self.update_relationships_display()
        self.update_data_type_combo()
        
        # Recreate data type frames in main tab
        if hasattr(self, 'data_sources_frame'):
            # Clear existing frames
            for widget in self.data_sources_frame.winfo_children():
                widget.destroy()
            
            # Recreate frames
            self.create_data_type_frames(self.data_sources_frame)
            
        # Refresh data types list in settings tab
        if hasattr(self, 'data_types_tree'):
            self.refresh_data_types_list()
            
    def refresh_data_types_list(self):
        """Refresh the data types treeview in settings tab"""
        if not hasattr(self, 'data_types_tree'):
            return
            
        # Clear existing items
        for item in self.data_types_tree.get_children():
            self.data_types_tree.delete(item)
            
        # Get all data types with their configurations
        data_types = self.relationship_system.get_data_types_dict()
        primary_type = self.relationship_system.get_primary_data_type()
        
        # Add each data type to the treeview
        for key, config in data_types.items():
            name = config.get('name', key.upper())
            short_name = config.get('short_name', key.upper())
            description = config.get('description', 'No description')
            
            # Determine status
            status = "Primary" if key == primary_type else "Secondary"
            if config.get('data'):
                status += " (Loaded)"
            else:
                status += " (Not Loaded)"
                
            # Insert into treeview
            self.data_types_tree.insert('', 'end', values=(key, name, short_name, description, status))
            
    def configure_selected_data_type(self):
        """Configure the selected data type in the settings tab"""
        selection = self.data_types_tree.selection()
        if not selection:
            messagebox.showwarning("Warning", "Please select a data type to configure")
            return
            
        # Get the selected item
        item = selection[0]
        data_type = self.data_types_tree.item(item, 'values')[0]
        
        # Open configuration dialog
        self.open_data_type_config_dialog(data_type)
        
    def remove_selected_data_type(self):
        """Remove the selected data type from the settings tab"""
        selection = self.data_types_tree.selection()
        if not selection:
            messagebox.showwarning("Warning", "Please select a data type to remove")
            return
            
        # Get the selected item
        item = selection[0]
        data_type = self.data_types_tree.item(item, 'values')[0]
        name = self.data_types_tree.item(item, 'values')[1]
        
        # Confirm deletion
        result = messagebox.askyesno("Confirm Removal", 
                                   f"Are you sure you want to remove the data type '{name}' ({data_type})?\n\n"
                                   "This will also remove all relationships involving this data type.")
        
        if result:
            try:
                # Remove the data type
                self.relationship_system.remove_data_type(data_type)
                
                # Refresh the GUI
                self.refresh_gui()
                
                messagebox.showinfo("Success", f"Data type '{name}' removed successfully")
                
            except Exception as e:
                messagebox.showerror("Error", f"Failed to remove data type: {e}") 