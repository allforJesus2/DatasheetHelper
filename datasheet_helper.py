print("DEBUG: Starting imports...")

print("DEBUG: Importing tkinter...")
import tkinter as tk
from tkinter import filedialog

print("DEBUG: Importing main_functions...")
from main_functions import *

print("DEBUG: Importing tkinter components...")
from tkinter import scrolledtext, Toplevel, Listbox, Button, Frame, MULTIPLE, Checkbutton, IntVar, Label, Scrollbar, Canvas, Entry, Spinbox, END, BOTH, LEFT, RIGHT, TOP, BOTTOM, X, Y, W, E, NW, SE, EW, messagebox, VERTICAL
from tkinter import ttk  # Import ttk for Combobox

print("DEBUG: Importing standard libraries...")
import json
import re
import os
# Set environment variable to handle OpenMP runtime conflict
os.environ['KMP_DUPLICATE_LIB_OK'] = 'TRUE'

print("DEBUG: Importing xlwings...")
import xlwings as xw

print("DEBUG: Importing Data_Extraction...")
from Data_Extraction import DatasheetExtractor

print("DEBUG: Importing xlsx_search...")
from xlsx_search import ExcelSearchApp

print("DEBUG: Importing edit_xlsx...")
from edit_xlsx import ExcelEditorApp

print("DEBUG: Importing pickle...")
import pickle

print("DEBUG: Importing tkinter.simpledialog...")
from tkinter.simpledialog import askstring

print("DEBUG: Importing excel_manager...")
from excel_manager import *

print("DEBUG: Importing openpyxl...")
import openpyxl

print("DEBUG: Importing excel_macro_viewer...")
from excel_macro_viewer import ExcelMacroViewer # Add this import

print("DEBUG: Importing excel_regex_search...")
from excel_regex_search import ExcelRegexSearchApp

# Semantic matcher will be imported lazily when needed

print("DEBUG: Importing threading...")
import threading
from datetime import datetime

print("DEBUG: Importing numpy...")
import numpy as np



print("DEBUG: All imports completed successfully!")


class CoordinateValue:
    def __init__(self, name, value, coordinate, sigfigs, input_units, output_units):
        self.name = name
        self.value = value
        self.coordinate = coordinate
        self.sigfigs = sigfigs
        self.input_units = input_units
        self.output_units = output_units

# Function to create the combined configuration dialog
def configure_source_data_dialog(parent, title, current_headers, file_path, current_selected_sheets, current_tolerance):
    dialog = Toplevel(parent)
    dialog.title(title)
    dialog.geometry("600x600") # Adjusted size
    dialog.transient(parent)
    dialog.grab_set()

    result = {
        "headers": current_headers.copy(), # Work on a copy
        "selected_sheets": current_selected_sheets,
        "tolerance": current_tolerance,
        "ok_pressed": False
    }

    # --- Header Configuration Section ---
    header_frame = ttk.LabelFrame(dialog, text="Headers")
    header_frame.pack(padx=10, pady=5, fill=X)

    header_list_frame = Frame(header_frame)
    header_list_frame.pack(pady=5, fill=X)
    header_listbox = Listbox(header_list_frame, height=6)
    header_listbox.pack(side=LEFT, fill=X, expand=True, padx=5)
    header_scrollbar = Scrollbar(header_list_frame, orient="vertical", command=header_listbox.yview)
    header_scrollbar.pack(side=RIGHT, fill=Y)
    header_listbox.config(yscrollcommand=header_scrollbar.set)

    for item in result["headers"]:
        header_listbox.insert(END, item)

    header_entry_frame = Frame(header_frame)
    header_entry_frame.pack(fill=X, padx=5, pady=(0, 5))
    header_entry = Entry(header_entry_frame)
    header_entry.pack(side=LEFT, fill=X, expand=True)

    def add_header():
        new_item = header_entry.get()
        if new_item and new_item not in result["headers"]:
            result["headers"].append(new_item)
            header_listbox.insert(END, new_item)
            header_entry.delete(0, END)
    header_entry.bind('<Return>', lambda event: add_header())

    def remove_header():
        selected_indices = header_listbox.curselection()
        if selected_indices:
            # Remove from listbox and result list in reverse order of index
            for index in sorted(selected_indices, reverse=True):
                item_to_remove = header_listbox.get(index)
                header_listbox.delete(index)
                if item_to_remove in result["headers"]:
                    result["headers"].remove(item_to_remove)

    header_button_frame = Frame(header_entry_frame)
    header_button_frame.pack(side=LEFT, padx=5)
    Button(header_button_frame, text="Add", command=add_header).pack(side=LEFT)
    Button(header_button_frame, text="Remove", command=remove_header).pack(side=LEFT, padx=5)

    # --- Sheet Selection Section ---
    sheet_frame = ttk.LabelFrame(dialog, text="Sheets to Process")
    sheet_frame.pack(padx=10, pady=5, fill=BOTH, expand=True)

    sheet_list_frame = Frame(sheet_frame)
    sheet_list_frame.pack(fill=BOTH, expand=True, padx=5, pady=5)

    sheet_canvas = Canvas(sheet_list_frame)
    sheet_scrollbar = Scrollbar(sheet_list_frame, orient="vertical", command=sheet_canvas.yview)
    sheet_scrollable_frame = Frame(sheet_canvas)

    sheet_scrollable_frame.bind(
        "<Configure>",
        lambda e: sheet_canvas.configure(scrollregion=sheet_canvas.bbox("all"))
    )

    sheet_canvas.create_window((0, 0), window=sheet_scrollable_frame, anchor=NW)
    sheet_canvas.configure(yscrollcommand=sheet_scrollbar.set)

    sheet_scrollbar.pack(side=RIGHT, fill=Y)
    sheet_canvas.pack(side=LEFT, fill=BOTH, expand=True)

    check_vars = {}
    all_sheet_names = []
    try:
        if file_path and os.path.exists(file_path):
            wb = openpyxl.load_workbook(file_path, read_only=True)
            all_sheet_names = wb.sheetnames
            wb.close()
        else:
            Label(sheet_scrollable_frame, text="Provide a valid file path first.").pack()
    except Exception as e:
        Label(sheet_scrollable_frame, text=f"Error loading sheets: {e}").pack()

    effective_selected_sheets = result["selected_sheets"] if result["selected_sheets"] is not None else all_sheet_names

    for sheet in all_sheet_names:
        var = IntVar(value=1 if sheet in effective_selected_sheets else 0)
        check_vars[sheet] = var
        Checkbutton(sheet_scrollable_frame, text=sheet, variable=var).pack(anchor=W, fill=X)

    # --- Tolerance Section ---
    tolerance_frame = ttk.LabelFrame(dialog, text="Header Detection Tolerance")
    tolerance_frame.pack(padx=10, pady=5, fill=X)
    Label(tolerance_frame, text="Max consecutive empty cells:").pack(side=LEFT, padx=5, pady=5)
    tolerance_spinbox = Spinbox(tolerance_frame, from_=0, to=10, width=5)
    tolerance_spinbox.delete(0, "end")
    tolerance_spinbox.insert(0, result["tolerance"])
    tolerance_spinbox.pack(side=LEFT, padx=5, pady=5)

    # --- Buttons Section ---
    button_frame = Frame(dialog)
    button_frame.pack(fill=X, padx=10, pady=10, side=BOTTOM)

    sheet_button_frame = Frame(button_frame) # Frame for sheet selection buttons
    sheet_button_frame.pack(side=LEFT)

    def on_select_all():
        for var in check_vars.values():
            var.set(1)

    def on_select_none():
        for var in check_vars.values():
            var.set(0)

    Button(sheet_button_frame, text="Select All Sheets", command=on_select_all).pack(side=LEFT, padx=5)
    Button(sheet_button_frame, text="Select None Sheets", command=on_select_none).pack(side=LEFT, padx=5)

    ok_cancel_frame = Frame(button_frame) # Frame for OK/Cancel buttons
    ok_cancel_frame.pack(side=RIGHT)

    def on_ok():
        result["selected_sheets"] = [sheet for sheet, var in check_vars.items() if var.get() == 1]
        try:
            result["tolerance"] = int(tolerance_spinbox.get())
        except ValueError:
            messagebox.showerror("Invalid Input", "Tolerance must be an integer.", parent=dialog)
            return # Keep dialog open
        result["ok_pressed"] = True
        dialog.destroy()

    def on_cancel():
        dialog.destroy()

    Button(ok_cancel_frame, text="OK", command=on_ok, width=10).pack(side=LEFT, padx=5)
    Button(ok_cancel_frame, text="Cancel", command=on_cancel, width=10).pack(side=LEFT, padx=5)

    parent.wait_window(dialog)
    return result if result["ok_pressed"] else None

class DatasheetGeneratorApp:

    # region Initialization

    def __init__(self, root):
        print("DEBUG: Starting DatasheetGeneratorApp initialization...")
        self.root = root
        print("DEBUG: Setting window title...")
        self.root.title("Datasheet Helper App")
        print("DEBUG: Creating ExcelManager...")
        self.excel_mgr = ExcelManager()
        print("DEBUG: Initializing basic attributes...")
        self.new_sheets = []
        self.halt_flag = False  # Flag to control halting of add_update_datasheets
        self.is_processing = False  # Flag to track if datasheet generation is running
        
        # Centralized data type definitions
        print("DEBUG: Setting up data type definitions...")
        self.data_types = {}
        
        # Initialize default settings
        print("DEBUG: Setting up default settings...")
        self.default_settings = {
            'transformation_code': 'int(x.split("-")[2])',
            'current_transform_data_type': None,  # Dynamic transform data type
            'current_transform_key': '',  # Dynamic transform key
            'tag_filters': [],
            'tag_cell_values': {},
            'datasheets': '',
            'source_sheet_name': 'TEMPLATE',
            'datasheet_coord': 'U8',
            'ds_str': 'DS-IA-',
            'rows_per_sheet': 1,
            'blank_cell_tolerance': 2,
            'sig_figs': 4, # Default significant figures
            'rounding_tolerance': 1e-2, # Default rounding tolerance
            'coordinate_conversions': {}, # Stores coordinate-specific unit conversions
            'coordinate_combinations': {} # Stores coordinate combinations (e.g., {'A1': {'combines': ['B1', 'C1'], 'operation': 'add'}})
        }
        
        # Add data type paths dynamically
        print("DEBUG: Adding data type paths...")
        for data_type, config in self.data_types.items():
            path_key = config.get('path_key', f'{data_type}_path')
            self.default_settings[path_key] = ''

        # Initialize all default settings as attributes
        print("DEBUG: Setting default setting attributes...")
        for param, value in self.default_settings.items():
            setattr(self, param, value)
        
        # Initialize transform-related attributes
        print("DEBUG: Setting transform attributes...")

        # Semantic similarity model
        print("DEBUG: Initializing semantic model attributes...")
        self.semantic_model = None
        self.model_loaded = False
        self.loading_model = False
        
        # Centralized Excel selection monitoring
        self.current_excel_selection = None
        self.coordinate_update_callbacks = []  # List of callback functions for each tab
        self.excel_monitoring_active = False

        print("DEBUG: About to create widgets...")
        self.create_widgets()
        print("DEBUG: DatasheetGeneratorApp initialization completed!")
        
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        # Start centralized Excel monitoring
        self.start_excel_monitoring()
    
    def get_data_type_config(self, data_type):
        """Get configuration for a specific data type"""
        return self.data_types.get(data_type, {})
    
    # region Centralized Excel Selection Monitoring
    
    def start_excel_monitoring(self):
        """Start the centralized Excel selection monitoring"""
        if not self.excel_monitoring_active:
            self.excel_monitoring_active = True
            self.monitor_excel_selection()
    
    def stop_excel_monitoring(self):
        """Stop the centralized Excel selection monitoring"""
        self.excel_monitoring_active = False
    
    def register_coordinate_callback(self, callback_func):
        """Register a callback function to be called when Excel selection changes"""
        if callback_func not in self.coordinate_update_callbacks:
            self.coordinate_update_callbacks.append(callback_func)
    
    def unregister_coordinate_callback(self, callback_func):
        """Unregister a callback function"""
        if callback_func in self.coordinate_update_callbacks:
            self.coordinate_update_callbacks.remove(callback_func)
    
    def cleanup_tab_callbacks(self, tab):
        """Clean up callbacks for a specific tab when it's destroyed"""
        if hasattr(tab, '_coordinate_callback'):
            self.unregister_coordinate_callback(tab._coordinate_callback)
            delattr(tab, '_coordinate_callback')
    
    def monitor_excel_selection(self):
        """Centralized Excel selection monitoring - runs once and updates all tabs"""
        if not self.excel_monitoring_active:
            return
            
        try:
            if self.datasheets:
                # Get current Excel selection
                full_selection = xw.apps.active.selection.address
                current_selection = full_selection.split(':')[0].replace('$', '')
                
                # Only update if selection has changed
                if current_selection != self.current_excel_selection:
                    self.current_excel_selection = current_selection
                    
                    # Call all registered callbacks
                    for callback in self.coordinate_update_callbacks[:]:  # Copy list to avoid modification during iteration
                        try:
                            callback(full_selection, current_selection)
                        except Exception as e:
                            print(f"Error in coordinate callback: {e}")
                            # Remove problematic callback
                            self.unregister_coordinate_callback(callback)
                            
        except Exception as e:
            # Excel might not be available or selection might be invalid
            pass
        
        # Schedule next check
        if self.excel_monitoring_active:
            self.root.after(200, self.monitor_excel_selection)
    
    # endregion
    
    def get_data_type_name(self, data_type):
        """Get the full name of a data type"""
        return data_type
    
    
    def get_data_type_headers(self, data_type):
        """Get the headers for a data type"""
        config = self.get_data_type_config(data_type)
        return config.get('headers', [])
    
    def get_data_type_source_sheet_name(self, data_type):
        """Get the source sheet name for a data type"""
        config = self.get_data_type_config(data_type)
        return config.get('source_sheet_name', 'TEMPLATE')
    
    def set_data_type_source_sheet_name(self, data_type, source_sheet_name):
        """Set the source sheet name for a data type"""
        if data_type in self.data_types:
            self.data_types[data_type]['source_sheet_name'] = source_sheet_name
    
    def set_data_type_headers(self, data_type, headers):
        """Set the headers for a data type"""
        if data_type in self.data_types:
            self.data_types[data_type]['headers'] = headers
    
    def get_data_type_data(self, data_type):
        """Get the data dictionary for a data type"""
        config = self.get_data_type_config(data_type)
        return config.get('data', {})
    
    def set_data_type_data(self, data_type, data):
        """Set the data dictionary for a data type"""
        if data_type in self.data_types:
            self.data_types[data_type]['data'] = data
            # Update the tab indicator to show data status
            self.update_tab_indicator(data_type)
    
    def get_data_type_coordinate_values(self, data_type):
        """Get coordinate values for a data type"""
        config = self.get_data_type_config(data_type)
        return config.get('coordinate_values', {})
    
    def set_data_type_coordinate_values(self, data_type, coordinate_values):
        """Set coordinate values for a data type"""
        if data_type in self.data_types:
            self.data_types[data_type]['coordinate_values'] = coordinate_values
    
    def get_data_type_selected_sheets(self, data_type):
        """Get selected sheets for a data type"""
        config = self.get_data_type_config(data_type)
        return config.get('selected_sheets', None)
    
    def set_data_type_selected_sheets(self, data_type, selected_sheets):
        """Set selected sheets for a data type"""
        if data_type in self.data_types:
            self.data_types[data_type]['selected_sheets'] = selected_sheets
    
    def get_data_type_path(self, data_type):
        """Get the file path for a data type"""
        config = self.get_data_type_config(data_type)
        path_key = config.get('path_key', f'{data_type}_path')
        return getattr(self, path_key, '')
    
    def set_data_type_path(self, data_type, path):
        """Set the file path for a data type"""
        config = self.get_data_type_config(data_type)
        path_key = config.get('path_key', f'{data_type}_path')
        setattr(self, path_key, path)
    
    def get_all_data_types(self):
        """Get all available data types"""
        return list(self.data_types.keys())
    
    def get_primary_data_type(self):
        """Get the primary data type (currently selected tab)"""
        if hasattr(self, 'data_sources_notebook'):
            try:
                # Get the currently selected tab index
                selected_index = self.data_sources_notebook.index(self.data_sources_notebook.select())
                # Get all data types in the same order as tabs
                data_types = list(self.get_all_data_types())
                if 0 <= selected_index < len(data_types):
                    return data_types[selected_index]
            except Exception as e:
                print(f"Error getting selected tab: {e}")
        
        # Fallback to first data type if no tab is selected or error occurs
        return list(self.data_types.keys())[0] if self.data_types else None
    
    def get_data_type_top_tag(self, data_type):
        """Get the top tag for a specific data type"""
        config = self.get_data_type_config(data_type)
        return config.get('top_tag', 'A1')
    
    def set_data_type_top_tag(self, data_type, top_tag):
        """Set the top tag for a specific data type"""
        if data_type in self.data_types:
            self.data_types[data_type]['top_tag'] = top_tag
    
    def on_tab_changed(self, event):
        """Handle tab change event to update primary data type"""
        primary_data_type = self.get_primary_data_type()
        if primary_data_type:
            config = self.get_data_type_config(primary_data_type)
            name = primary_data_type
            print(f"Primary data type changed to: {name} ({primary_data_type})")
    
    def add_data_type_from_entry(self, event=None):
        """Add a new data type from the text entry box"""
        data_type_id = self.add_data_type_entry.get().strip()
        
        # Validation
        if not data_type_id:
            messagebox.showerror("Error", "Data Type ID is required")
            return
        
        # Check if data type already exists
        if data_type_id in self.data_types:
            messagebox.showerror("Error", f"Data type '{data_type_id}' already exists")
            return
        
        # Create the new data type configuration
        new_config = {
            'headers': [],
            'coordinate_values': {},
            'selected_sheets': None,
            'path_key': f'{data_type_id}_path',
            'top_tag': 'A1',  # Default top tag
            'source_sheet_name': 'TEMPLATE',  # Default source sheet name
            'data': {}  # Data will be stored as nested dictionaries here
        }
        
        # Add to data types
        self.add_data_type(data_type_id, new_config)
        
        # Clear the entry box
        self.add_data_type_entry.delete(0, tk.END)
        
        # Refresh the interface
        self.refresh_data_sources_notebook()
        
        messagebox.showinfo("Success", f"Data type '{data_type_id}' created successfully!")

    def add_data_type(self, data_type, config):
        """Add a new data type configuration"""
        # Ensure the config has all required fields with defaults
        default_config = {
            'headers': [],
            'coordinate_values': {},
            'selected_sheets': None,
            'top_tag': 'A1',
            'source_sheet_name': 'TEMPLATE',
            'data': {}
        }
        default_config.update(config)
        
        # Add to the centralized data types dictionary
        self.data_types[data_type] = default_config
        
        # Add the path to default settings if it doesn't exist
        path_key = default_config.get('path_key', f'{data_type}_path')
        if path_key not in self.default_settings:
            self.default_settings[path_key] = ''
            setattr(self, path_key, '')  # Still need this for file path variables
        
        # Refresh the GUI to include the new data type
        if hasattr(self, 'root') and self.root:
            self.refresh_data_type_frames()
    

    
    def create_data_type_frames(self, parent):
        """Dynamically create UI frames for all data types"""
        print("DEBUG: Starting create_data_type_frames...")
        # Add a label for the data sources section
        sources_label = tk.Label(parent, text="DATA SOURCES", font=("Arial", 10, "bold"), fg="green")
        sources_label.pack(anchor=tk.W, padx=5, pady=(5,0))
        
        print("DEBUG: Creating frames for each data type...")
        for data_type in self.get_all_data_types():
            config = self.get_data_type_config(data_type)
            name = data_type
            path_key = config.get('path_key', f'{data_type}_path')
            
            # Create frame for this data type
            frame = tk.Frame(parent)
            frame.pack(fill=tk.X, pady=5)
            
            # Label
            label = tk.Label(frame, text=f"{name} (Source)", width=30)
            label.pack(side=tk.LEFT, padx=5)
            
            # Entry
            entry = tk.Entry(frame)
            entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
            self.entries.append((entry, path_key))
            
            # Buttons frame
            buttons = tk.Frame(frame)
            buttons.pack(side=tk.RIGHT)
            
            # Browse button
            tk.Button(buttons, text="Browse",
                      command=lambda e=entry, p=path_key: self.browse(e, p)).pack(side=tk.LEFT, padx=2)
            
            # Open button
            tk.Button(buttons, text="Open",
                      command=lambda e=entry: self.open_file(e)).pack(side=tk.LEFT, padx=2)
            
            # Configure button
            tk.Button(buttons, text="Configure",
                      command=lambda n=name: self.configure(n)).pack(side=tk.LEFT, padx=2)
            
            # Generate button
            tk.Button(buttons, text="Generate",
                      command=lambda dt=data_type: self.generate_data_type(dt)).pack(side=tk.LEFT, padx=2)
            
            # View button
            tk.Button(buttons, text="View",
                      command=lambda n=name: self.view_data(n)).pack(side=tk.LEFT, padx=2)
            
            # Store reference to frame for potential updates
            setattr(self, f'{data_type}_frame', frame)
        print("DEBUG: create_data_type_frames completed!")
    
    def create_data_sources_notebook(self, parent):
        """Create a notebook with tabs for each data source, each containing nested configuration tabs"""
        print("DEBUG: Starting create_data_sources_notebook...")
        
        # Add a label and button for the data sources section
        sources_frame = ttk.Frame(parent)
        sources_frame.pack(fill="x", padx=5, pady=(5,0))
        
        sources_label = tk.Label(sources_frame, text="DATA SOURCES", font=("Arial", 10, "bold"), fg="green")
        sources_label.pack(side="left")
        
        # Add text entry and button to create new data type
        add_data_type_frame = ttk.Frame(sources_frame)
        add_data_type_frame.pack(side="right")
        
        self.add_data_type_entry = ttk.Entry(add_data_type_frame, width=15)
        self.add_data_type_entry.pack(side="left", padx=(0, 5))
        self.add_data_type_entry.bind('<Return>', self.add_data_type_from_entry)
        self.add_data_type_entry.insert(0, "Enter data type...")
        self.add_data_type_entry.bind('<FocusIn>', lambda e: self.add_data_type_entry.delete(0, tk.END) if self.add_data_type_entry.get() == "Enter data type..." else None)
        
        add_data_type_btn = ttk.Button(add_data_type_frame, text="+ Add Data Type", 
                                      command=self.add_data_type_from_entry, width=15)
        add_data_type_btn.pack(side="left")
        
        # Create the main data sources notebook
        self.data_sources_notebook = ttk.Notebook(parent)
        self.data_sources_notebook.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # Create tabs for each data type
        for data_type in self.get_all_data_types():
            config = self.get_data_type_config(data_type)
            name = data_type
            
            # Create main tab for this data type
            data_source_tab = ttk.Frame(self.data_sources_notebook)
            self.data_sources_notebook.add(data_source_tab, text=data_type)
            
            # Create the data source tab content
            self.create_data_source_tab_content(data_source_tab, data_type, config)
        
        # Update all tab indicators after creating tabs
        self.update_all_tab_indicators()
        
        # Bind tab change event to update primary data type
        self.data_sources_notebook.bind("<<NotebookTabChanged>>", self.on_tab_changed)
        
        # Bind right-click event for context menu
        self.data_sources_notebook.bind("<Button-3>", self.show_tab_context_menu)
        
        print("DEBUG: create_data_sources_notebook completed!")
    
    def update_tab_indicator(self, data_type):
        """Update the tab text to show data status with a check mark"""
        if not hasattr(self, 'data_sources_notebook'):
            return
        
        # Find the tab index for this data type
        tab_index = None
        for i in range(self.data_sources_notebook.index("end")):
            tab_text = self.data_sources_notebook.tab(i, "text")
            # Remove any existing check mark to get the base data type
            base_data_type = tab_text.replace(" ✓", "")
            if base_data_type == data_type:
                tab_index = i
                break
        
        if tab_index is not None:
            # Check if data is populated
            data = self.get_data_type_data(data_type)
            has_data = bool(data and len(data) > 0)
            
            # Update tab text with or without check mark
            if has_data:
                new_text = f"{data_type} ✓"
            else:
                new_text = data_type
            
            self.data_sources_notebook.tab(tab_index, text=new_text)
    
    def update_all_tab_indicators(self):
        """Update indicators for all data type tabs"""
        for data_type in self.get_all_data_types():
            self.update_tab_indicator(data_type)
    
    def show_tab_context_menu(self, event):
        """Show context menu when right-clicking on a tab"""
        # Get the tab index at the click position
        tab_index = self.data_sources_notebook.index(f"@{event.x},{event.y}")
        if tab_index is None:
            return
        
        # Get the data type from the tab text (remove check mark if present)
        tab_text = self.data_sources_notebook.tab(tab_index, "text")
        data_type = tab_text.replace(" ✓", "")  # Remove check mark to get base data type
        
        # Create context menu
        context_menu = tk.Menu(self.root, tearoff=0)
        context_menu.add_command(label="Rename Data Type", 
                               command=lambda: self.rename_data_type(data_type))
        context_menu.add_separator()
        context_menu.add_command(label="Delete Data Type", 
                               command=lambda: self.delete_data_type(data_type))
        
        # Show the context menu
        try:
            context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            context_menu.grab_release()
    
    def rename_data_type(self, old_data_type):
        """Rename a data type"""
        # Ask for new name
        new_data_type = tk.simpledialog.askstring("Rename Data Type", 
                                                 f"Enter new name for '{old_data_type}':",
                                                 initialvalue=old_data_type)
        
        if not new_data_type or new_data_type.strip() == "":
            return
        
        new_data_type = new_data_type.strip()
        
        # Check if new name already exists
        if new_data_type in self.data_types:
            messagebox.showerror("Error", f"Data type '{new_data_type}' already exists")
            return
        
        # Check if it's the same name
        if new_data_type == old_data_type:
            return
        
        # Get the old configuration
        old_config = self.data_types[old_data_type].copy()
        
        # Create new data type with the new name
        self.data_types[new_data_type] = old_config
        
        # Update path_key in the config
        old_path_key = old_config.get('path_key', f'{old_data_type}_path')
        new_path_key = f'{new_data_type}_path'
        self.data_types[new_data_type]['path_key'] = new_path_key
        
        # Update default settings path
        if old_path_key in self.default_settings:
            self.default_settings[new_path_key] = self.default_settings[old_path_key]
            del self.default_settings[old_path_key]
        
        # Update instance attributes
        if hasattr(self, old_path_key):
            setattr(self, new_path_key, getattr(self, old_path_key))
            delattr(self, old_path_key)
        
        # Remove old data type
        del self.data_types[old_data_type]
        
        # Refresh the interface
        self.refresh_data_sources_notebook()
        
        messagebox.showinfo("Success", f"Data type renamed from '{old_data_type}' to '{new_data_type}'")
    
    def delete_data_type(self, data_type):
        """Delete a data type"""
        # Confirm deletion
        if not messagebox.askyesno("Confirm Delete", 
                                  f"Are you sure you want to delete data type '{data_type}'?\n\nThis action cannot be undone."):
            return
        
        # Remove from data types
        if data_type in self.data_types:
            del self.data_types[data_type]
        
        # Remove from default settings
        path_key = f'{data_type}_path'
        if path_key in self.default_settings:
            del self.default_settings[path_key]
        
        # Remove instance attribute
        if hasattr(self, path_key):
            delattr(self, path_key)
        
        # Refresh the interface
        self.refresh_data_sources_notebook()
        
        messagebox.showinfo("Success", f"Data type '{data_type}' deleted")
    
    def refresh_data_sources_notebook(self):
        """Add new data type tabs and remove old ones to match current data types"""
        if not hasattr(self, 'data_sources_notebook'):
            return
        
        # Get current data types
        current_data_types = set(self.get_all_data_types())
        
        # Get all existing tab texts and their indices
        existing_tabs = {}
        for i in range(self.data_sources_notebook.index("end")):
            tab_text = self.data_sources_notebook.tab(i, "text")
            existing_tabs[tab_text] = i
        
        # Remove tabs that no longer exist in data types
        tabs_to_remove = []
        for tab_text, tab_index in existing_tabs.items():
            if tab_text not in current_data_types:
                tabs_to_remove.append(tab_index)
        
        # Remove tabs in reverse order to maintain indices
        for tab_index in sorted(tabs_to_remove, reverse=True):
            # Find the tab text for this index
            tab_text = None
            for text, idx in existing_tabs.items():
                if idx == tab_index:
                    tab_text = text
                    break
            self.data_sources_notebook.forget(tab_index)
            if tab_text:
                print(f"Removed tab for data type: {tab_text}")
        
        # Add tabs for any new data types that don't exist yet
        for data_type in current_data_types:
            if data_type not in existing_tabs:
                config = self.get_data_type_config(data_type)
                name = data_type
                
                # Create main tab for this data type
                data_source_tab = ttk.Frame(self.data_sources_notebook)
                self.data_sources_notebook.add(data_source_tab, text=data_type)
                
                # Create the data source tab content
                self.create_data_source_tab_content(data_source_tab, data_type, config)
                
                print(f"Added new tab for data type: {data_type}")
        
        # Update all tab indicators
        self.update_all_tab_indicators()
        
        # Switch to the first tab if any exist
        if self.data_sources_notebook.index("end") > 0:
            self.data_sources_notebook.select(0)
    
    def create_data_source_tab_content(self, parent, data_type, config):
        """Create the content for a data source tab, including file selection and nested configuration tabs"""
        name = data_type
        path_key = config.get('path_key', f'{data_type}_path')
        
        # Create file selection frame at the top
        file_frame = ttk.LabelFrame(parent, text=f"{name} File Selection")
        file_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # File selection row
        file_row = tk.Frame(file_frame)
        file_row.pack(fill=tk.X, padx=5, pady=5)
        
        # Label
        label = tk.Label(file_row, text=f"{name} File:", width=20)
        label.pack(side=tk.LEFT, padx=5)
        
        # Entry
        entry = tk.Entry(file_row)
        entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        self.entries.append((entry, path_key))
        
        # Buttons frame
        buttons = tk.Frame(file_row)
        buttons.pack(side=tk.RIGHT)
        
        # Browse button
        tk.Button(buttons, text="Browse",
                  command=lambda e=entry, p=path_key: self.browse(e, p)).pack(side=tk.LEFT, padx=2)
        
        # Open button
        tk.Button(buttons, text="Open",
                  command=lambda e=entry: self.open_file(e)).pack(side=tk.LEFT, padx=2)
        
        # Configure button
        tk.Button(buttons, text="Configure",
                  command=lambda n=name: self.configure(n)).pack(side=tk.LEFT, padx=2)
        
        # Generate button
        tk.Button(buttons, text="Generate",
                  command=lambda dt=data_type: self.generate_data_type(dt)).pack(side=tk.LEFT, padx=2)
        
        # View button
        tk.Button(buttons, text="View",
                  command=lambda n=name: self.view_data(n)).pack(side=tk.LEFT, padx=2)
        
        # Create nested notebook for configuration tabs (Coordinates, Filters, Transform)
        config_notebook = ttk.Notebook(parent)
        config_notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # Create the three configuration tabs
        coordinates_tab = ttk.Frame(config_notebook)
        filters_tab = ttk.Frame(config_notebook)
        transform_tab = ttk.Frame(config_notebook)
        
        config_notebook.add(coordinates_tab, text='Coordinates')
        config_notebook.add(filters_tab, text='Filters')
        config_notebook.add(transform_tab, text='Transform')
        
        # Create tab content
        self.create_coordinates_tab(coordinates_tab, data_type)
        self.create_filters_tab(filters_tab)
        self.create_transform_tab(transform_tab)
        
        # Create source sheet name frame below the coordinate tabs
        source_sheet_frame = ttk.LabelFrame(parent, text=f"{name} Source Sheet Configuration")
        source_sheet_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # Source sheet name row
        source_sheet_row = tk.Frame(source_sheet_frame)
        source_sheet_row.pack(fill=tk.X, padx=5, pady=5)
        
        # Label
        source_sheet_label = tk.Label(source_sheet_row, text="Source Sheet Name:", width=20)
        source_sheet_label.pack(side=tk.LEFT, padx=5)
        
        # Entry for source sheet name
        source_sheet_entry = tk.Entry(source_sheet_row)
        source_sheet_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
        # Set initial value
        current_source_sheet = self.get_data_type_source_sheet_name(data_type)
        source_sheet_entry.insert(0, current_source_sheet)
        
        # Bind change event to update the data type configuration
        def update_source_sheet_name(event=None):
            new_source_sheet = source_sheet_entry.get().strip()
            if new_source_sheet:
                self.set_data_type_source_sheet_name(data_type, new_source_sheet)
        
        source_sheet_entry.bind('<KeyRelease>', update_source_sheet_name)
        source_sheet_entry.bind('<FocusOut>', update_source_sheet_name)
        
        # Store reference to the config notebook for this data type
        setattr(self, f'{data_type}_config_notebook', config_notebook)
    
    def generate_data_type(self, data_type):
        """Generic method to generate data for any data type"""
        # Use the generic approach for all data types - truly extensible
        self.generate_generic_data(data_type)
    
    def generate_generic_data(self, data_type):
        """Generate data for any data type using the centralized system"""

        config = self.get_data_type_config(data_type)
        name = data_type
        path_key = config.get('path_key', f'{data_type}_path')
        
        # Get the file path
        file_path = getattr(self, path_key, '')
        if not file_path:
            messagebox.showwarning("Warning", f"No file path set for {name}")
            return
        
        # Determine file type and process accordingly
        file_extension = os.path.splitext(file_path)[1].lower()
        
        if file_extension == '.json':
            # Process JSON file
            try:
                from main_functions import load_dict_from_json
                data = load_dict_from_json(file_path)
                self.set_data_type_data(data_type, data)
                self.refresh_tab_content()
                messagebox.showinfo("Success", f"Successfully loaded {name} data from JSON file.")
                print(f"Loaded {name} data from JSON: {file_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load JSON file:\n{str(e)}")
                
        elif file_extension in ['.xlsx', '.xls']:
            # Show custom dialog for Excel processing
            choice = self.show_excel_processing_dialog(data_type, file_path)
            
            if choice == "datasheet":
                # Process as datasheet
                self.load_data_type_from_datasheet(data_type, file_path)
            elif choice == "index":
                # Process as index (generate dictionary)
                self.process_excel_as_index(file_path, data_type)
            # If choice is None (Cancel), do nothing
        else:
            messagebox.showerror("Error", f"Unsupported file type for {name}: {file_extension}")
    
    def process_excel_as_index(self, file_path, data_type):
        """Process Excel file as index (generate dictionary from headers)"""
        
        config = self.get_data_type_config(data_type)
        name = data_type
        
        # Get headers and selected sheets
        headers = self.get_data_type_headers(data_type)
        selected_sheets = self.get_data_type_selected_sheets(data_type)
        
        # Check if configuration is missing
        if not headers or not selected_sheets:
            # Show full configuration dialog
            result = configure_source_data_dialog(
                self.root, 
                f"Configure {name} Data Source",
                headers or [], 
                file_path, 
                selected_sheets or [], 
                self.blank_cell_tolerance
            )
            
            if result and result["ok_pressed"]:
                # Update configuration with user's choices
                self.set_data_type_headers(data_type, result["headers"])
                self.set_data_type_selected_sheets(data_type, result["selected_sheets"])
                self.blank_cell_tolerance = result["tolerance"]
                
                # Use the updated configuration
                headers = result["headers"]
                selected_sheets = result["selected_sheets"]
                print(f"{name} configuration updated.")
            else:
                print(f"{name} configuration cancelled.")
                return
        
        # Generate data using the centralized data type pattern
        data = generate_dictionary_from_xlsx(file_path, headers,
                                          parent=self.root, selected_sheets=selected_sheets,
                                          max_empty_allowed=self.blank_cell_tolerance)
        
        if data is not None:
            self.set_data_type_data(data_type, data)
            show_nested_dict_analysis(data)
            print(f"Generated {name}")
            self.refresh_tab_content()
        else:
            print(f"{name} generation cancelled or failed.")
    
    def refresh_data_type_frames(self):
        """Refresh all data type frames when new data types are added"""
        # Remove existing data type frames
        for data_type in self.get_all_data_types():
            frame_name = f'{data_type}_frame'
            if hasattr(self, frame_name):
                frame = getattr(self, frame_name)
                frame.destroy()
                delattr(self, frame_name)
        
        # Recreate frames
        main_frame = self.root.winfo_children()[0]  # Get the main frame
        self.create_data_type_frames(main_frame)

        #self.root.protocol("WM_DELETE_WINDOW", self.on_closing)


    def create_widgets(self):
        print("DEBUG: Starting create_widgets...")
        # Create menu bar
        print("DEBUG: Creating menu bar...")
        self.menu_bar = tk.Menu(self.root)
        self.root.config(menu=self.menu_bar)

        # Create Commands menu
        print("DEBUG: Creating Commands menu...")
        self.command_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label="Commands", menu=self.command_menu)

        # Add menu items
        print("DEBUG: Setting up menu commands...")
        menu_commands = [
            ("Load Settings", self.load_settings),
            ("Save Settings", self.save_settings),
            ("Run xlsx search app", self.open_excel_search_app),
            ("Populate Headers on Datasheets", self.open_edit_xlsx),
            ("View Coordinate Value Data", self.display_coordinate_values),
            ("Delete newly added datasheets", self.delete_added_sheets),
            ("Rebuild tabs", self.refresh_tab_content),
            ("Delete Certain Sheets by Prefix", self.delete_sheets_by_prefix),
            ("Excel Macros", self.open_excel_macros_window),
            ("Excel Regex Search App", self.open_excel_regex_search_app),
            ("Semantic Matcher", self.open_semantic_matcher),
            ("Release Excel", self.release_excel_connection),
            ("Stop Datasheet Generation", self.set_halt_flag)
        ]
        
        # Add dynamic menu items for each data type
        print("DEBUG: Adding dynamic menu items for data types...")
        for data_type in self.get_all_data_types():
            config = self.get_data_type_config(data_type)
            name = data_type
            
            # Data type specific menu items removed - now handled by browse button

        print("DEBUG: Adding menu commands to menu...")
        for label, command in menu_commands:
            self.command_menu.add_command(label=label, command=command)

        # Create main container frame
        print("DEBUG: Creating main container frame...")
        main_frame = tk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        self.entries = []  # Store entries for later reference

        # Create data sources notebook with tabs for each data type
        print("DEBUG: Creating data sources notebook...")
        self.create_data_sources_notebook(main_frame)

        # Add a visual separator between data sources and destination
        print("DEBUG: Creating separator and destination container...")
        separator_frame = tk.Frame(main_frame, height=2, bg='gray')
        separator_frame.pack(fill=tk.X, pady=10)
        separator_frame.pack_propagate(False)

        # Create destination container with distinct styling
        destination_container = tk.Frame(main_frame, relief=tk.RAISED, borderwidth=2)
        destination_container.pack(fill=tk.X, pady=5, padx=5)

        # Destination label to indicate this is the output
        dest_label = tk.Label(destination_container, text="DESTINATION", font=("Arial", 10, "bold"), fg="blue")
        dest_label.pack(anchor=tk.W, padx=5, pady=(5,0))

        # Datasheets Row (single destination)
        ds_frame = tk.Frame(destination_container)
        ds_frame.pack(fill=tk.X, pady=5)

        ds_label = tk.Label(ds_frame, text="Datasheets (Destination)", width=30)
        ds_label.pack(side=tk.LEFT, padx=5)

        ds_entry = tk.Entry(ds_frame)
        ds_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        self.entries.append((ds_entry, "datasheets"))

        ds_buttons = tk.Frame(ds_frame)
        ds_buttons.pack(side=tk.RIGHT)

        tk.Button(ds_buttons, text="Browse",
                  command=lambda: self.browse_destination(ds_entry)).pack(side=tk.LEFT, padx=2)
        tk.Button(ds_buttons, text="Open",
                  command=self.configure_ds).pack(side=tk.LEFT, padx=2)
        tk.Button(ds_buttons, text="Datasheet Configuration",
                  command=self.open_datasheet_config_window).pack(side=tk.LEFT, padx=2)


        # Color coding method and buttons on same line
        controls_frame = tk.Frame(destination_container)
        controls_frame.pack(fill=tk.X, pady=5)

        color_label = tk.Label(controls_frame, text="Color Coding Method:", width=20)
        color_label.pack(side=tk.LEFT, padx=5)

        self.color_coding_var = tk.StringVar(value="new_red_old_green")
        color_dropdown = ttk.Combobox(controls_frame, textvariable=self.color_coding_var, 
                                     values=["None (Black)", "new_red_old_green", "new_red"], 
                                     state="readonly", width=20)
        color_dropdown.pack(side=tk.LEFT, padx=5)

        # Add some spacing
        tk.Frame(controls_frame, width=20).pack(side=tk.LEFT)

        self.generate_button = tk.Button(controls_frame, text="Add/Update",
                  command=self.add_datasheets, font=("Arial", 10, "bold"), 
                  bg="green", fg="white", padx=20, pady=5)
        self.generate_button.pack(side=tk.LEFT, padx=5)
        
        self.stop_button = tk.Button(controls_frame, text="Stop",
                  command=self.set_halt_flag, bg="red", fg="white", state="disabled",
                  font=("Arial", 10, "bold"), padx=20, pady=5)
        self.stop_button.pack(side=tk.LEFT, padx=5)

        # Add status label
        self.status_label = tk.Label(destination_container, text="Ready", fg="black", font=("Arial", 9))
        self.status_label.pack(anchor=tk.W, padx=5, pady=(5,0))

        print("DEBUG: create_widgets completed!")

    # endregion
    # region Tab Creation

    def init_excel(self):
        """Initialize Excel only when needed"""
        if self.datasheets:
            try:
                # Check if workbook reference is still valid
                if self.excel_mgr.wb and self.excel_mgr.wb.name:
                    return
            except:
                # Workbook was closed, reset references
                self.excel_mgr.wb = None
                self.excel_mgr.app = None
            
            print("DEBUG: About to open workbook with ExcelManager...")
            self.excel_mgr.open_workbook(self.datasheets)
            print("DEBUG: Workbook opened successfully with ExcelManager")

    def create_coordinates_tab(self, tab, data_type=None, sheet_names=None):
        print(f"DEBUG: Starting create_coordinates_tab for data_type: {data_type}...")

        def get_combo_values():
            """Get combo values for the specific data type"""
            if data_type:
                data = self.get_data_type_data(data_type)
                values = []
                if data:
                    # Get keys from the first entry
                    first_entry = list(data.values())[0]
                    if isinstance(first_entry, dict):
                        values = list(first_entry.keys())
                return {data_type: values}
            else:
                # Fallback for backward compatibility
                combo_values = {}
                for dt in self.get_all_data_types():
                    data = self.get_data_type_data(dt)
                    values = []
                    if data:
                        # Get keys from the first entry
                        first_entry = list(data.values())[0]
                        if isinstance(first_entry, dict):
                            values = list(first_entry.keys())
                    combo_values[dt] = values
            return combo_values

        def add_coordinate(data_type, entry, combo, listbox):
            coord = entry.get()
            value = combo.get()
            if coord and value:
                coordinate_values = self.get_data_type_coordinate_values(data_type)
                coordinate_values[coord] = value
                self.set_data_type_coordinate_values(data_type, coordinate_values)
                
                update_listboxes()

        def remove_coordinate(data_type, listbox):
            selected = listbox.curselection()
            if selected:
                idx = selected[0]
                coord = listbox.get(idx).split(':')[0]
                coordinate_values = self.get_data_type_coordinate_values(data_type)
                if coord in coordinate_values:
                    del coordinate_values[coord]
                    self.set_data_type_coordinate_values(data_type, coordinate_values)
                update_listboxes()

        def clear_coordinates(data_type, listbox):
            self.set_data_type_coordinate_values(data_type, {})
            update_listboxes()

        def update_coordinate_display(full_selection, current_selection):
            """Callback function for centralized Excel selection monitoring"""
            try:
                # Update coordinate entry
                entry_var.set(current_selection)
                
                # Update region display
                try:
                    clean_selection = full_selection.replace('$', '')
                    
                    if ',' in clean_selection:
                        # Non-contiguous selection (multiple ranges separated by commas)
                        ranges = clean_selection.split(',')
                        range_descriptions = []
                        for i, range_addr in enumerate(ranges):
                            if ':' in range_addr:
                                start_cell, end_cell = range_addr.split(':')
                                range_descriptions.append(f"Range{i+1}: {start_cell}:{end_cell}")
                            else:
                                range_descriptions.append(f"Cell{i+1}: {range_addr}")
                        
                        # Create detailed display
                        if len(ranges) <= 3:
                            # Show all ranges if 3 or fewer
                            region_var.set(f"Selected: {', '.join(range_descriptions)}")
                        else:
                            # Show count if more than 3 ranges
                            region_var.set(f"Selected: {len(ranges)} Non-contiguous Ranges ({ranges[0]}, {ranges[1]}, ...)")
                            
                    elif ':' in clean_selection:
                        # Single contiguous range
                        start_cell, end_cell = clean_selection.split(':')
                        region_var.set(f"Selected: Range {start_cell}:{end_cell}")
                    else:
                        # Single cell selected
                        region_var.set(f"Selected: Single Cell {clean_selection}")
                except Exception as e:
                    region_var.set("Selected: Single Cell")
                
                # Update cell values display
                try:
                    sheet = xw.apps.active.books.active.sheets.active
                    above_value, left_value = self.get_cell_values_above_and_left(sheet, current_selection)
                    
                    # Simple truncation to prevent layout issues
                    above_text = str(above_value)[:15] + "..." if above_value and len(str(above_value)) > 15 else (str(above_value) if above_value else "—")
                    left_text = str(left_value)[:15] + "..." if left_value and len(str(left_value)) > 15 else (str(left_value) if left_value else "—")
                    
                    cell_values_var.set(f"Above: {above_text} | Left: {left_text}")
                except Exception as e:
                    cell_values_var.set("Above: — | Left: —")
                    
            except Exception as e:
                print(f"Error updating coordinate display: {e}")
        
        # Register this tab's callback with the centralized monitoring
        self.register_coordinate_callback(update_coordinate_display)
        
        # Store callback reference for cleanup
        tab._coordinate_callback = update_coordinate_display

        def update_listbox(data_type, listbox):
            """Update a specific data type's listbox using centralized system"""
            try:
                # Check if the widget still exists
                if not listbox.winfo_exists():
                    return
                listbox.delete(0, tk.END)
                coordinate_values = self.get_data_type_coordinate_values(data_type)
                for key, value in coordinate_values.items():
                    # Add placeholder for conversion details
                    coord_display = f"{key}: {value}"
                    if key in self.coordinate_conversions:
                        conv = self.coordinate_conversions[key]
                        coord_display += f" [{conv.get('in_unit', '?')}->{conv.get('out_unit', '?')}]"
                    if key in self.coordinate_combinations:
                        combo = self.coordinate_combinations[key]
                        coord_display += f" [Combines: {', '.join(combo.get('combines', []))} ({combo.get('operation', 'add')})]"
                    listbox.insert(tk.END, coord_display)
            except tk.TclError:
                # Widget was destroyed, skip updating
                print(f"Warning: Could not update {data_type} listbox - widget may have been destroyed")
                return

        def update_listboxes():
            """Update all listboxes for all data types"""
            for data_type in self.get_all_data_types():
                listbox_name = f"{data_type}_listbox"
                if hasattr(self, listbox_name):
                    try:
                        listbox = getattr(self, listbox_name)
                        # Check if the widget still exists before updating
                        if listbox.winfo_exists():
                            update_listbox(data_type, listbox)
                    except tk.TclError:
                        # Widget was destroyed, skip updating
                        print(f"Warning: Could not update {listbox_name} - widget may have been destroyed")
                        continue

        def reinitialize():
            #init_excel()
            combo_values = get_combo_values()
            for data_type in self.get_all_data_types():
                combo_name = f"{data_type}_combo"
                if hasattr(self, combo_name):
                    combo = getattr(self, combo_name)
                    combo['values'] = combo_values.get(data_type, [])
            update_listboxes()

        # Initial Excel setup if path exists
        #init_excel()
        combo_values = get_combo_values()

        # Create main container with left and right sections
        main_container = ttk.Frame(tab)
        main_container.pack(fill="both", expand=True, padx=5, pady=5)
        
        # Left side - Main coordinate controls
        left_frame = ttk.Frame(main_container)
        left_frame.pack(side="left", fill="both", expand=True, padx=(0, 5))
        
        # Right side - Semantic mapping container
        right_frame = ttk.LabelFrame(main_container, text="Semantic Mapping", padding="5")
        right_frame.pack(side="right", fill="y", padx=(5, 0))
        right_frame.configure(width=300)  # Fixed width for right panel
        
        # UI Setup with Top Tag entry for the specific data type
        top_frame = ttk.Frame(left_frame)
        top_frame.pack(fill="x", pady=(0, 5))

        # Top Tag entry (only show for the specific data type)
        if data_type:
            current_top_tag = self.get_data_type_top_tag(data_type)
            config = self.get_data_type_config(data_type)
            name = data_type
            
            top_tag_frame = ttk.Frame(top_frame)
            top_tag_frame.pack(fill="x", pady=(0, 5))
            
            ttk.Label(top_tag_frame, text=f"Top Tag for {name}:").pack(side="left")
            top_tag_entry = ttk.Entry(top_tag_frame, width=10)
            top_tag_entry.pack(side="left", padx=(5, 0))
            top_tag_entry.insert(0, current_top_tag)
            
            def update_top_tag():
                new_top_tag = top_tag_entry.get().strip()
                if new_top_tag:
                    self.set_data_type_top_tag(data_type, new_top_tag)
                    print(f"Updated {name} top tag to: {new_top_tag}")
            
            ttk.Button(top_tag_frame, text="Update", command=update_top_tag).pack(side="left", padx=(5, 0))
            
            # Add help button for top tag
            def show_top_tag_help():
                help_text = """Top Tag - Key Coordinate

This is the Excel cell coordinate (e.g., A1, I12) where the tag identifier is placed in each datasheet.

WHAT IT DOES:
• Places the dictionary key (tag identifier) at this coordinate
• Serves as the anchor point for all tag-related operations
• Used to identify existing tags when updating sheets
• Used to place new tags when creating sheets
• Calculates row offsets for positioning other data

HOW IT WORKS WITH COORDINATE-VALUE DATA:
The Top Tag and Coordinate-Value Data work together:

1. TOP TAG: Determines WHERE the tag goes
   - Places the dictionary key (e.g., "TAG-001") at the specified coordinate
   - Example: If Top Tag = "A1", then "TAG-001" goes in cell A1

2. COORDINATE-VALUE DATA: Determines WHAT data goes where
   - Contains mapping of coordinates to values for each tag
   - Example: {"B1": "Line 1", "C1": "PID-001", "D1": "100.5"}

COMPLETE PROCESS FLOW:
1. System reads your data: {"TAG-001": {"B1": "Line 1", "C1": "PID-001"}}
2. Top Tag places "TAG-001" at the key coordinate (e.g., A1)
3. Coordinate-Value Data places "Line 1" at B1, "PID-001" at C1, etc.

EXAMPLE WITH DETAILS:
If Top Tag = "A1" and you have data like:
{"TAG-001": {"B1": "Line 1", "C1": "PID-001", "D1": "100.5"}}

Result in Excel:
• Cell A1: "TAG-001" (the tag identifier from dictionary key)
• Cell B1: "Line 1" (from coordinate-value data)
• Cell C1: "PID-001" (from coordinate-value data)
• Cell D1: "100.5" (from coordinate-value data)

AUTOMATIC SETTING:
• The Top Tag is automatically set to the first coordinate you add for the primary data type
• You can manually change it if needed

IMPORTANT NOTES:
• The Top Tag is the ANCHOR POINT for all tag operations
• All other data positioning is calculated relative to this coordinate
• When updating existing sheets, the system looks for tags at this coordinate
• When creating new sheets, tags are placed at this coordinate"""
                
                help_window = tk.Toplevel(tab)
                help_window.title("Top Tag Help")
                help_window.geometry("500x400")
                help_window.transient(tab)
                help_window.grab_set()
                
                # Create scrolled text widget
                import tkinter.scrolledtext as scrolledtext
                text_widget = scrolledtext.ScrolledText(help_window, wrap=tk.WORD, padx=10, pady=10)
                text_widget.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
                text_widget.insert(tk.END, help_text)
                text_widget.configure(state='disabled')
                
                # Close button
                ttk.Button(help_window, text="Close", command=help_window.destroy).pack(pady=10)
            
            help_btn = ttk.Button(top_tag_frame, text="?", width=3, command=show_top_tag_help)
            help_btn.pack(side="left", padx=(5, 0))

        # Coordinate entry row with cell values and selection info
        coord_frame = ttk.Frame(top_frame)
        coord_frame.pack(fill="x", pady=(0, 5))
        
        coord_label = ttk.Label(coord_frame, text="Enter Key Coordinate:")
        coord_label.pack(side="left")

        entry_var = tk.StringVar()
        coord_entry = ttk.Entry(coord_frame, textvariable=entry_var)
        coord_entry.pack(side="left", fill="x", expand=True, padx=(5, 0))
        
        # Cell values display on same line
        cell_values_var = tk.StringVar()
        cell_values_var.set("Above: — | Left: —")
        cell_values_label = ttk.Label(coord_frame, textvariable=cell_values_var, 
                                     font=("Arial", 9), foreground="gray")
        cell_values_label.pack(side="left", padx=(10, 0))

        # Region selection display on same line
        region_var = tk.StringVar()
        region_var.set("Selected: Single Cell")
        region_label = ttk.Label(coord_frame, textvariable=region_var, 
                                font=("Arial", 9), foreground="blue")
        region_label.pack(side="left", padx=(10, 0))

        # Main content frame (moved to left frame)
        content_frame = ttk.Frame(left_frame)
        content_frame.pack(fill="both", expand=True)

        # Create frame for the specific data type (or all if data_type is None for backward compatibility)
        data_types_to_process = [data_type] if data_type else self.get_all_data_types()
        
        for current_data_type in data_types_to_process:
            config = self.get_data_type_config(current_data_type)
            
            # Create frame for this data type
            data_frame = ttk.LabelFrame(content_frame, text=f"{current_data_type} Coordinates")
            data_frame.pack(side="left", fill="both", expand=True, padx=5)

            # Controls frame
            controls = ttk.Frame(data_frame)
            controls.pack(fill="x", padx=5, pady=5)

            # Label
            label = ttk.Label(controls, text=f"Select {current_data_type} Value:")
            label.pack(side="left")

            # Combo box
            combo = ttk.Combobox(controls, values=combo_values.get(current_data_type, []), state="readonly")
            combo.pack(side="left", fill="x", expand=True, padx=5)

            # Button frame
            btn_frame = ttk.Frame(data_frame)
            btn_frame.pack(fill="x", padx=5)

            # Listbox
            listbox = tk.Listbox(data_frame, height=8)
            listbox.pack(fill="both", expand=True, padx=5, pady=5)

            # Buttons
            ttk.Button(btn_frame, text=f"Add to {current_data_type}",
                       command=lambda dt=current_data_type, e=coord_entry, c=combo, l=listbox: add_coordinate(dt, e, c, l)).pack(side="left", padx=2)
            ttk.Button(btn_frame, text="Remove",
                       command=lambda dt=current_data_type, l=listbox: remove_coordinate(dt, l)).pack(side="left", padx=2)
            ttk.Button(btn_frame, text="Clear All",
                       command=lambda dt=current_data_type, l=listbox: clear_coordinates(dt, l)).pack(side="left", padx=2)
            
            # AutoMap button with min score entry
            def automap_coordinate(dt=current_data_type, combo_box=combo):                    # Get min score from right panel
                min_score = float(min_score_var_right.get().strip())

                full_selection = xw.apps.active.selection.address
                print(f"DEBUG: Full selection: {full_selection}")
                # Load semantic model if not already loaded
                if not self.model_loaded and not self.loading_model:
                    self.load_semantic_model_async()
                    return
                
                # Check selection type
                clean_selection = full_selection.replace('$', '')
                
                if ',' in clean_selection:
                    # Non-contiguous selection - process multiple ranges
                    print(f"DEBUG: Non-contiguous selection: {clean_selection}")
                    self.automap_noncontiguous(dt, full_selection, combo_box, coord_entry, listbox, min_score)
                elif ':' in clean_selection:
                    print(f"DEBUG: Single contiguous range: {clean_selection}")
                    # Single contiguous range - iterate through cells
                    self.automap_range(dt, full_selection, combo_box, coord_entry, listbox, min_score)
                else:
                    # Single cell - perform mapping
                    print(f"DEBUG: Single cell: {clean_selection}")
                    current_coord = coord_entry.get().strip()
                    if current_coord:
                        best_match, score, header = self.auto_map_coordinate_semantic(dt, current_coord, min_score)
                        if best_match:
                            combo_box.set(best_match)
                            add_coordinate(dt, coord_entry, combo_box, listbox)
                            header_display = f"'{header}'" if header else "None"
                            messagebox.showinfo("AutoMap Result", 
                                f"Successfully mapped {current_coord} to '{best_match}'\nSimilarity Score: {score:.3f}\nHeader: {header_display}")
                        else:
                            header_display = f"'{header}'" if header else "None"
                            messagebox.showinfo("AutoMap Result", 
                                f"No match found for {current_coord}\nBest Score: {score:.3f} (below threshold {min_score})\nHeader: {header_display}")
            
            # AutoMap button
            ttk.Button(btn_frame, text="AutoMap",
                       command=lambda dt=current_data_type, cb=combo: automap_coordinate(dt, cb)).pack(side="left", padx=2)

            # Store references for later use
            setattr(self, f"{current_data_type}_combo", combo)
            setattr(self, f"{current_data_type}_listbox", listbox)

        
        # Semantic mapping controls (moved to right frame)
        # Status indicator
        self.semantic_status_label = ttk.Label(right_frame, text="Semantic model: Not loaded", foreground="red")
        self.semantic_status_label.pack(anchor="w", pady=(0, 5))
        
        # Load model button
        def load_semantic_model():
            if not self.model_loaded and not self.loading_model:
                self.load_semantic_model_async()
                self.semantic_status_label.config(text="Semantic model: Loading...", foreground="orange")
        
        ttk.Button(right_frame, text="Load Semantic Model", 
                   command=load_semantic_model).pack(anchor="w", pady=(0, 10))
        
        # Help text
        help_text = "Auto Map: Uses AI to find the best matching field from your data based on text above/left of selected cell"
        help_label = ttk.Label(right_frame, text=help_text, font=("Arial", 9), 
                 foreground="gray", wraplength=280, justify="left")
        help_label.pack(anchor="w", pady=(0, 10))
        
        # AutoMap settings in right panel
        automap_frame = ttk.LabelFrame(right_frame, text="AutoMap Settings", padding="5")
        automap_frame.pack(fill="x", pady=(0, 10))
        
        # Min score setting
        score_frame = ttk.Frame(automap_frame)
        score_frame.pack(fill="x", pady=(0, 5))
        
        ttk.Label(score_frame, text="Min Score:").pack(side="left")
        min_score_var_right = tk.StringVar(value="0.3")
        min_score_entry_right = ttk.Entry(score_frame, textvariable=min_score_var_right, width=8)
        min_score_entry_right.pack(side="right")
        
        # Min score is now only in the right panel

        # --- Context Menu Setup ---
        # Create context menu if it doesn't exist
        if not hasattr(self, 'coord_context_menu'):
            self.coord_context_menu = tk.Menu(self.root, tearoff=0)
            self.coord_context_menu.add_command(label="Change Source Key", command=self.open_change_source_key_dialog)
            self.coord_context_menu.add_separator()
            self.coord_context_menu.add_command(label="Add/Edit Conversion", command=self.open_conversion_dialog)
            self.coord_context_menu.add_command(label="Remove Conversion", command=self.remove_conversion)
            self.coord_context_menu.add_command(label="Combine", command=self.open_combination_dialog)
            self.coord_context_menu.add_command(label="Remove Combination", command=self.remove_combination)
            self.coord_context_menu.add_separator()
            self.coord_context_menu.add_command(label="Cancel")

        if not hasattr(self, 'selected_coord_for_context'):
            self.selected_coord_for_context = None

        def show_coord_context_menu(event, listbox_widget, data_type):
            # Select the item under the cursor
            clicked_index = listbox_widget.nearest(event.y)
            listbox_widget.selection_clear(0, tk.END)
            listbox_widget.selection_set(clicked_index)
            listbox_widget.activate(clicked_index)

            selected_text = listbox_widget.get(clicked_index)
            # Extract coordinate (part before ':')
            self.selected_coord_for_context = selected_text.split(':')[0].strip()

            # Check if conversion exists and enable/disable "Remove Conversion"
            if self.selected_coord_for_context in self.coordinate_conversions:
                self.coord_context_menu.entryconfig("Remove Conversion", state="normal")
            else:
                self.coord_context_menu.entryconfig("Remove Conversion", state="disabled")

            # Check if combination exists and enable/disable "Remove Combination"
            if self.selected_coord_for_context in self.coordinate_combinations:
                self.coord_context_menu.entryconfig("Remove Combination", state="normal")
            else:
                self.coord_context_menu.entryconfig("Remove Combination", state="disabled")

            # Popup the menu
            try:
                self.coord_context_menu.tk_popup(event.x_root, event.y_root)
            finally:
                self.coord_context_menu.grab_release()

        # Bind right-click to this specific listbox
        listbox.bind("<Button-3>", lambda event, l=listbox, dt=data_type: show_coord_context_menu(event, l, dt))
        # ---

        update_listboxes()

        # Save/Load coordinate maps functionality
        def save_coordinate_maps():
            """Save all coordinate maps to a JSON file"""
            try:
                file_path = filedialog.asksaveasfilename(
                    title="Save Coordinate Maps",
                    defaultextension=".json",
                    filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
                )
                if file_path:
                    coordinate_data = {}
                    for data_type in self.get_all_data_types():
                        coordinate_data[data_type] = self.get_data_type_coordinate_values(data_type)
                    
                    # Also save coordinate conversions and combinations
                    coordinate_data['conversions'] = self.coordinate_conversions
                    coordinate_data['combinations'] = self.coordinate_combinations
                    
                    with open(file_path, 'w') as f:
                        json.dump(coordinate_data, f, indent=2)
                    
                    messagebox.showinfo("Success", f"Coordinate maps saved to {file_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save coordinate maps: {e}")
        
        def load_coordinate_maps():
            """Load coordinate maps from a JSON file"""
            try:
                file_path = filedialog.askopenfilename(
                    title="Load Coordinate Maps",
                    filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
                )
                if file_path:
                    with open(file_path, 'r') as f:
                        coordinate_data = json.load(f)
                    
                    # Load coordinate values for each data type
                    for data_type in self.get_all_data_types():
                        if data_type in coordinate_data:
                            self.set_data_type_coordinate_values(data_type, coordinate_data[data_type])
                    
                    # Load coordinate conversions and combinations
                    if 'conversions' in coordinate_data:
                        self.coordinate_conversions = coordinate_data['conversions']
                    if 'combinations' in coordinate_data:
                        self.coordinate_combinations = coordinate_data['combinations']
                    
                    # Update the UI
                    update_listboxes()
                    messagebox.showinfo("Success", f"Coordinate maps loaded from {file_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load coordinate maps: {e}")
        
        # Add Save/Load buttons to the coordinates tab
        save_load_frame = ttk.Frame(tab)
        save_load_frame.pack(fill="x", padx=10, pady=5)
        
        ttk.Button(save_load_frame, text="Save Coordinate Maps", 
                   command=save_coordinate_maps).pack(side="left", padx=5)
        ttk.Button(save_load_frame, text="Load Coordinate Maps", 
                   command=load_coordinate_maps).pack(side="left", padx=5)

        # Expose reinitialize method for external calls
        tab.reinitialize = reinitialize
        return tab

    def create_filters_tab(self, tab):
        print("DEBUG: Starting create_filters_tab...")
        # Add filter frame

        # Add explanatory label for filter functionality
        filter_info_label = ttk.Label(tab, text="Can use comma to include multiple filter terms")
        filter_info_label.pack(anchor="w", padx=5, pady=(2, 0))
        add_frame = ttk.Frame(tab)
        add_frame.pack(fill="x", padx=5, pady=2)

        filters_entries = []
        # Get combo values using centralized system from all data types
        combo_values = []
        for data_type in self.get_all_data_types():
            data = self.get_data_type_data(data_type)
            if data:
                # Get keys from the first entry
                first_entry = list(data.values())[0]
                if isinstance(first_entry, dict):
                    combo_values.extend(list(first_entry.keys()))
        
        # Remove duplicates while preserving order
        combo_values = list(dict.fromkeys(combo_values))

        def get_filter_values_for_header(header):
            """Get unique values for a specific header from all data types"""
            unique_values = set()
            for data_type in self.get_all_data_types():
                data = self.get_data_type_data(data_type)
                if data:
                    for entry_key, entry_data in data.items():
                        if isinstance(entry_data, dict) and header in entry_data and entry_data[header] is not None:
                            unique_values.add(str(entry_data[header]))
            return sorted(list(unique_values))

        def add_filter_row(name='', filter_value=''):
            new_row = len(filters_entries) + 1
            name_label = tk.Label(content_frame, text=f"Index Key {new_row}:")
            name_label.grid(row=new_row, column=0)
            name_entry = ttk.Combobox(content_frame, values=combo_values)
            name_entry.grid(row=new_row, column=1)
            name_entry.set(name)

            filter_label = tk.Label(content_frame, text=f"Filter {new_row}:")
            filter_label.grid(row=new_row, column=2)
            
            # Get filter values for the selected header
            filter_values = get_filter_values_for_header(name) if name else []
            filter_entry = ttk.Combobox(content_frame, values=filter_values)
            filter_entry.grid(row=new_row, column=3)
            filter_entry.set(filter_value)

            # Update filter values when header changes
            def on_header_change(event=None):
                selected_header = name_entry.get()
                new_filter_values = get_filter_values_for_header(selected_header)
                filter_entry['values'] = new_filter_values
                # Clear current selection if it's not valid for new header
                if filter_entry.get() not in new_filter_values:
                    filter_entry.set('')
            
            name_entry.bind('<<ComboboxSelected>>', on_header_change)
            name_entry.bind('<KeyRelease>', on_header_change)

            filters_entries.append((name_entry, filter_entry))

        def save_filters():
            self.tag_filters.clear()
            for name_entry, filter_entry in filters_entries:
                name = name_entry.get()
                filter_value = filter_entry.get()
                if name and filter_value:
                    self.tag_filters.append([name, filter_value])

        content_frame = ttk.Frame(tab)
        content_frame.pack(fill="both", expand=True, padx=5, pady=5)

        button_frame = ttk.Frame(tab)
        button_frame.pack(fill="x", padx=5, pady=5)

        add_button = ttk.Button(button_frame, text="Add Filter", command=lambda: add_filter_row())
        add_button.pack(side="left", padx=5)

        save_button = ttk.Button(button_frame, text="Save", command=save_filters)
        save_button.pack(side="right", padx=5)

        for name, filter_value in self.tag_filters:
            add_filter_row(name, filter_value)

        add_filter_row()  # Add empty row

    def create_transform_tab(self, tab):
        print("DEBUG: Starting create_transform_tab...")
        # Source data type selector
        data_type_frame = ttk.Frame(tab)
        data_type_frame.pack(fill="x", padx=5, pady=2)

        ttk.Label(data_type_frame, text="Source Data Type:").pack(side="left")

        # Get available data types for combo
        data_type_values = [self.get_data_type_name(dt) for dt in self.get_all_data_types()]
        data_type_combo = ttk.Combobox(data_type_frame, values=data_type_values)
        data_type_combo.pack(side="left", fill="x", expand=True, padx=5)
        
        # Set default to first data type
        if data_type_values:
            data_type_combo.set(data_type_values[0])

        # Source key selector
        key_frame = ttk.Frame(tab)
        key_frame.pack(fill="x", padx=5, pady=2)

        ttk.Label(key_frame, text="Source Key:").pack(side="left")

        key_combo = ttk.Combobox(key_frame, values=[])
        key_combo.pack(side="left", fill="x", expand=True, padx=5)

        # Update key combo when data type changes
        def update_key_combo(*args):
            selected_data_type_name = data_type_combo.get()
            data_type = self.get_data_type_by_name(selected_data_type_name)
            if data_type:
                data = self.get_data_type_data(data_type)
                if data:
                    # Get keys from the first entry
                    first_entry = list(data.values())[0]
                    if isinstance(first_entry, dict):
                        key_values = list(first_entry.keys())
                        key_combo['values'] = key_values
                        if key_values:
                            key_combo.set(key_values[0])

        data_type_combo.bind('<<ComboboxSelected>>', update_key_combo)
        update_key_combo()  # Initialize

        # Transformation code entry
        code_frame = ttk.Frame(tab)
        code_frame.pack(fill="x", padx=5, pady=5)

        ttk.Label(code_frame, text="Key, x = Index Value:").pack(side="left")
        code_entry = ttk.Entry(code_frame)
        code_entry.pack(side="left", fill="x", expand=True, padx=5)
        code_entry.insert(0, self.transformation_code)

        # Add test section
        test_frame = ttk.LabelFrame(tab, text="Test Transformation")
        test_frame.pack(fill="x", padx=5, pady=5)
        
        test_input_frame = ttk.Frame(test_frame)
        test_input_frame.pack(fill="x", padx=5, pady=5)
        
        ttk.Label(test_input_frame, text="Test Input Value:").pack(side="left")
        test_input = ttk.Entry(test_input_frame)
        test_input.pack(side="left", fill="x", expand=True, padx=5)
        
        test_result_frame = ttk.Frame(test_frame)
        test_result_frame.pack(fill="x", padx=5, pady=5)
        
        ttk.Label(test_result_frame, text="Result:").pack(side="left")
        result_var = tk.StringVar()
        result_label = ttk.Label(test_result_frame, textvariable=result_var)
        result_label.pack(side="left", padx=5)
        
        def test_transform():
            input_value = test_input.get()
            if input_value:
                try:
                    transformation = code_entry.get()
                    from main_functions import translate
                    result = translate(input_value, transformation)
                    result_var.set(str(result))
                except Exception as e:
                    result_var.set(f"Error: {str(e)}")
            else:
                result_var.set("Please enter a test value")
        
        ttk.Button(test_frame, text="Test", command=test_transform).pack(pady=5)

        def save_transform():
            selected_data_type_name = data_type_combo.get()
            data_type = self.get_data_type_by_name(selected_data_type_name)
            self.current_transform_data_type = data_type
            self.current_transform_key = key_combo.get()
            self.transformation_code = code_entry.get()
            print(f"Transform saved - Data Type: {selected_data_type_name}, Key: {self.current_transform_key}, Code: {self.transformation_code}")

        ttk.Button(tab, text="Save", command=save_transform).pack(pady=5)

    def refresh_tab_content(self, force_rebuild=False):
        """Refresh all configuration tabs within each data source tab"""
        print("DEBUG: Starting refresh_tab_content...")
        data_types = self.get_all_data_types()
        print(f"DEBUG: Found {len(data_types)} data types: {data_types}")
        
        # Load sheet names once for all tabs to avoid multiple file access
        print("DEBUG: Loading sheet names once for all tabs...")
        sheet_names = self.get_sheet_names()
        print(f"DEBUG: Loaded {len(sheet_names)} sheet names for all tabs")
        
        # Refresh all configuration tabs within each data source tab
        for data_type in data_types:
            print(f"DEBUG: Refreshing tabs for data_type: {data_type}")
            config_notebook_name = f'{data_type}_config_notebook'
            if hasattr(self, config_notebook_name):
                print(f"DEBUG: Found config notebook: {config_notebook_name}")
                config_notebook = getattr(self, config_notebook_name)
                
                # Only rebuild if forced or if tabs don't exist
                existing_tabs = config_notebook.winfo_children()
                if force_rebuild or len(existing_tabs) == 0:
                    print(f"DEBUG: Rebuilding tabs for {data_type}")
                    # Recreate all tabs with updated data
                    print(f"DEBUG: Destroying existing tabs for {data_type}")
                    for tab in existing_tabs:
                        tab.destroy()

                    print(f"DEBUG: Creating new tabs for {data_type}")
                    coordinates_tab = ttk.Frame(config_notebook)
                    filters_tab = ttk.Frame(config_notebook)
                    transform_tab = ttk.Frame(config_notebook)

                    config_notebook.add(coordinates_tab, text='Coordinates')
                    config_notebook.add(filters_tab, text='Filters')
                    config_notebook.add(transform_tab, text='Transform')

                    print(f"DEBUG: Creating coordinates tab for {data_type}")
                    self.create_coordinates_tab(coordinates_tab, data_type, sheet_names)
                    print(f"DEBUG: Creating filters tab for {data_type}")
                    self.create_filters_tab(filters_tab)
                    print(f"DEBUG: Creating transform tab for {data_type}")
                    self.create_transform_tab(transform_tab)
                    print(f"DEBUG: Completed tabs for {data_type}")
                else:
                    print(f"DEBUG: Skipping rebuild for {data_type} - tabs already exist")
            else:
                print(f"DEBUG: No config notebook found for {data_type}")
        
        print("DEBUG: refresh_tab_content completed")

    def open_datasheet_config_window(self):
        """Open a new window with datasheet configuration options"""
        # Create new window
        config_window = tk.Toplevel(self.root)
        config_window.title("Datasheet Configuration")
        config_window.geometry("770x350")
        config_window.transient(self.root)
        config_window.grab_set()
        
        # Create main frame with padding
        main_frame = ttk.Frame(config_window, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Create scrollable frame
        canvas = tk.Canvas(main_frame)
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Pack canvas and scrollbar
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Create the datasheet configuration content in the scrollable frame
        self.create_datasheet_config_content(scrollable_frame, config_window)
        
        # Center the window
        config_window.update_idletasks()
        x = (config_window.winfo_screenwidth() // 2) - (config_window.winfo_width() // 2)
        y = (config_window.winfo_screenheight() // 2) - (config_window.winfo_height() // 2)
        config_window.geometry(f"+{x}+{y}")

    def create_datasheet_config_content(self, parent, window):
        """Create the datasheet configuration content for the popup window"""
        # Import scrolledtext for help window
        import tkinter.scrolledtext as scrolledtext
        
        fields = [
            ("Datasheet Coordinate", "datasheet_coord", ttk.Entry),
            ("Datasheet Prefix used to identify sheets names that should be updated and fill DS numbers\n(If blank, we will look for existing tags in all sheets starting from the top tag)\nThis could be problematic if there is a cover sheet we try pulling tags from which might cause issues with accessing sheets that dont exist", "ds_str", ttk.Entry),
            ("Rows per Sheet", "rows_per_sheet", ttk.Entry)
        ]

        entries = {}
        for i, (label, attr, widget_type) in enumerate(fields):
            ttk.Label(parent, text=label).pack(anchor="w", padx=5, pady=(5, 0))
            
            w = widget_type(parent)
            w.pack(fill="x", padx=5, pady=(0, 5))
            
            if widget_type == ttk.Combobox:
                # Use provided sheet names or get them if not provided
                if sheet_names is None:
                    print(f"DEBUG: Getting sheet names for combobox widget (attr: {attr})")
                    sheet_names = self.get_sheet_names()
                    print(f"DEBUG: Got {len(sheet_names)} sheet names for combobox")
                else:
                    print(f"DEBUG: Using provided sheet names for combobox widget (attr: {attr})")
                    print(f"DEBUG: Using {len(sheet_names)} provided sheet names for combobox")
                
                w['values'] = sheet_names
                # Set current value if it exists, otherwise use first sheet name
                current_value = getattr(self, attr, '')
                if current_value and current_value in sheet_names:
                    w.set(current_value)
                elif sheet_names:
                    w.set(sheet_names[0])
                print(f"DEBUG: Combobox configured with {len(sheet_names)} values")
            else:
                # For regular entry widgets, insert current value
                if hasattr(self, attr):
                    w.insert(0, getattr(self, attr))
            
            entries[attr] = w

        # Add Sig Figs and Tolerance fields
        ttk.Label(parent, text="Significant Figures for Rounding").pack(anchor="w", padx=5, pady=(5, 0))
        sig_figs_entry = ttk.Entry(parent)
        sig_figs_entry.pack(fill="x", padx=5, pady=(0, 5))
        sig_figs_entry.insert(0, getattr(self, 'sig_figs'))
        entries['sig_figs'] = sig_figs_entry

        ttk.Label(parent, text="Rounding Tolerance (e.g., 1e-2)").pack(anchor="w", padx=5, pady=(5, 0))
        tolerance_entry = ttk.Entry(parent)
        tolerance_entry.pack(fill="x", padx=5, pady=(0, 5))
        tolerance_entry.insert(0, getattr(self, 'rounding_tolerance'))
        entries['rounding_tolerance'] = tolerance_entry

        def save_datasheet_settings():
            for attr, entry in entries.items():
                if attr == 'rows_per_sheet':
                    try:
                        setattr(self, attr, int(entry.get()))
                    except ValueError:
                        setattr(self, attr, 0)
                elif attr == 'sig_figs': # Handle sig_figs
                    try:
                        setattr(self, attr, int(entry.get()))
                    except ValueError:
                        setattr(self, attr, 4) # Default if invalid
                        entry.delete(0, tk.END)
                        entry.insert(0, '4')
                        messagebox.showwarning("Invalid Input", "Significant figures must be an integer. Using default value 4.")
                elif attr == 'rounding_tolerance': # Handle rounding_tolerance
                    try:
                        setattr(self, attr, float(entry.get()))
                    except ValueError:
                        setattr(self, attr, 1e-2) # Default if invalid
                        entry.delete(0, tk.END)
                        entry.insert(0, '1e-2')
                        messagebox.showwarning("Invalid Input", "Rounding tolerance must be a number. Using default value 1e-2.")
                else:
                    setattr(self, attr, entry.get())
            print("Datasheet settings saved.")
            messagebox.showinfo("Success", "Datasheet configuration saved successfully!")
            window.destroy()

        # Button frame
        button_frame = ttk.Frame(parent)
        button_frame.pack(fill="x", padx=5, pady=10)
        
        save_btn = ttk.Button(button_frame, text="Save", command=save_datasheet_settings)
        save_btn.pack(side="left", padx=(0, 5))
        
        cancel_btn = ttk.Button(button_frame, text="Cancel", command=window.destroy)
        cancel_btn.pack(side="left")

    # endregion

    # region Configuration Windows


    def set_tag_filters(self):
        view_window = tk.Toplevel(self.root)
        view_window.title("Set Tag Filters (Comma for OR). ReGenerate Coordinates if necessary")

        filters_entries = []

        # Get combo values using centralized system from all data types
        combo_values = []
        for data_type in self.get_all_data_types():
            data = self.get_data_type_data(data_type)
            if data:
                for key, value in data.items():
                    combo_values.extend(list(value.keys()))
                    print(list(combo_values))
                    break
        
        # Remove duplicates while preserving order
        combo_values = list(dict.fromkeys(combo_values))

        def add_filter_row(name='', filter_value=''):
            new_row = len(filters_entries) + 1

            name_label = tk.Label(view_window, text=f"Index Key {new_row}:")
            name_label.grid(row=new_row, column=0)
            name_entry = ttk.Combobox(view_window, values=combo_values)
            name_entry.grid(row=new_row, column=1)
            name_entry.set(name)  # Prepopulate with existing name

            filter_label = tk.Label(view_window, text=f"Filter {new_row}:")
            filter_label.grid(row=new_row, column=2)
            filter_entry = tk.Entry(view_window)
            filter_entry.grid(row=new_row, column=3)
            filter_entry.insert(0, filter_value)  # Prepopulate with existing filter_value

            filters_entries.append((name_entry, filter_entry))

        def save_filters():
            self.tag_filters.clear()  # Clear self.tag_filters to update with new values
            for name_entry, filter_entry in filters_entries:
                name = name_entry.get()
                filter_value = filter_entry.get()
                if name and filter_value:
                    self.tag_filters.append([name, filter_value])

            # For demonstration, you may print or use the self.tag_filters list here
            print("Saved Tag Filters:")
            print(self.tag_filters)

            # Here, you might perform any required action with self.tag_filters

        add_button = tk.Button(view_window, text="Add New", command=add_filter_row)
        add_button.grid(row=0, column=0, columnspan=2, sticky='ew', padx=5, pady=5)

        save_button = tk.Button(view_window, text="Save", command=save_filters)
        save_button.grid(row=0, column=2, columnspan=2, sticky='ew', padx=5, pady=5)

        # Populate initial rows with existing tag filters from self.tag_filters
        for name, filter_value in self.tag_filters:
            add_filter_row(name, filter_value)

        # Add an empty row at the end
        add_filter_row()

        # Configure row and column weights to make them expandable
        for i in range(4):  # Assuming 4 rows in the layout (adjust if needed)
            view_window.grid_columnconfigure(i, weight=1)

        view_window.mainloop()

    # endregion

    # region Data Generation and Processing


    def assign_value_coordinate_to_tag(self):
        """Legacy method - now calls the dynamic version"""
        self.assign_value_coordinate_to_tag_dynamic()
    
    def assign_value_coordinate_to_tag_dynamic(self):
        """Dynamic version that works with any data types"""
        print("Generating Coordinate-Value Data")
        self.tag_cell_values = {}  # 'a1':'LINE', 'a2':'PID' ...
        
        def process_coordinate_value(coordinate, raw_value):
            """Helper function to process a coordinate value with conversions and combinations"""
            if raw_value is None:
                return None
                
            # Apply unit conversion first if it exists
            if coordinate in self.coordinate_conversions:
                conv_details = self.coordinate_conversions[coordinate]
                formula_str = conv_details.get('formula')
                if formula_str:
                    try:
                        # Attempt conversion only if value is numeric-like
                        numeric_value = float(str(raw_value).strip().replace(',', ''))
                        conversion_func = eval(formula_str)
                        converted_value = conversion_func(numeric_value)
                        print(f"  Converted {coordinate}: {raw_value} -> {converted_value}")
                        raw_value = converted_value
                    except Exception as conv_e:
                        print(f"  Conversion error for {coordinate} ({raw_value}): {conv_e}, using original value.")
            
            return raw_value
        
        def apply_combinations(data):
            """Apply coordinate combinations to the data dictionary"""
            for coord, combo_info in self.coordinate_combinations.items():
                if coord in data:  # Skip if the target coordinate already has a value
                    continue
                    
                combines = combo_info.get('combines', [])
                operation = combo_info.get('operation', 'add')
                
                # Get all the values to combine
                values_to_combine = []
                for combine_coord in combines:
                    if combine_coord in data and data[combine_coord] is not None:
                        values_to_combine.append(data[combine_coord])
                
                if not values_to_combine:
                    print(f"  Warning: No valid values found for combination {coord}")
                    continue
                
                try:
                    if operation == 'add':
                        # Convert to numeric and add
                        numeric_values = [float(str(v).strip().replace(',', '')) for v in values_to_combine]
                        result = sum(numeric_values)
                    elif operation == 'subtract':
                        # Convert to numeric and subtract (first value minus the rest)
                        numeric_values = [float(str(v).strip().replace(',', '')) for v in values_to_combine]
                        result = numeric_values[0] - sum(numeric_values[1:])
                    elif operation == 'multiply':
                        # Convert to numeric and multiply
                        numeric_values = [float(str(v).strip().replace(',', '')) for v in values_to_combine]
                        result = 1
                        for v in numeric_values:
                            result *= v
                    elif operation == 'divide':
                        # Convert to numeric and divide (first value divided by the rest)
                        numeric_values = [float(str(v).strip().replace(',', '')) for v in values_to_combine]
                        result = numeric_values[0]
                        for v in numeric_values[1:]:
                            if v != 0:
                                result /= v
                            else:
                                print(f"  Warning: Division by zero in combination {coord}")
                                result = None
                                break
                    elif operation == 'concatenate':
                        # Concatenate as strings
                        result = ''.join(str(v) for v in values_to_combine)
                    else:
                        print(f"  Warning: Unknown operation '{operation}' for combination {coord}")
                        continue
                    
                    if result is not None:
                        data[coord] = result
                        print(f"  Combined {coord}: {values_to_combine} ({operation}) -> {result}")
                        
                except (ValueError, TypeError) as e:
                    print(f"  Error processing combination {coord}: {e}")
                    continue
        
        # Get the primary data type (explicitly configured) - truly dynamic
        primary_data_type = self.get_primary_data_type()
        primary_data = self.get_data_type_data(primary_data_type)
        
        if not primary_data:
            print(f"No data available for primary data type: {primary_data_type}")
            return
            
        for tag in primary_data:
            if tag:
                # filter out
                continue_flag = False
                for header, filter_key in self.tag_filters:
                    print('tag', tag)
                    print('header', header)
                    # Split the filter key on commas to get multiple acceptable values
                    acceptable_values = [value.strip() for value in filter_key.split(',')]

                    # If the tag's value for this header isn't in our acceptable values, filter it out
                    if primary_data[tag][header] not in acceptable_values:
                        continue_flag = True
                        break

                if continue_flag:
                    continue

                # Process coordinates for the primary data type only
                data = {}
                
                # Process only the primary data type coordinates
                primary_data_type = self.get_primary_data_type()
                data_type_data = self.get_data_type_data(primary_data_type)
                coordinate_values = self.get_data_type_coordinate_values(primary_data_type)
                
                for coordinate, value in coordinate_values.items():
                    print(f'{primary_data_type} tag {tag}, value {value}, coord {coordinate}')
                    raw_value = data_type_data[tag].get(value) # Use .get() for safety
                    processed_value = process_coordinate_value(coordinate, raw_value)
                    data[coordinate] = processed_value
                    if processed_value is None:
                        print(f"  Warning: Key '{value}' not found in {primary_data_type} for tag '{tag}'. Skipping coordinate '{coordinate}'.")

                # Apply combinations to the data
                apply_combinations(data)

                self.tag_cell_values[tag] = data

        print("Coordinate Values generated:", self.tag_cell_values)
    
    def assign_value_coordinate_to_tag_simple_dynamic(self):
        """Simplified dynamic version that works with any data types"""
        print("Generating Coordinate-Value Data (Dynamic)")
        self.tag_cell_values = {}
        
        # Get the primary data type (explicitly configured)
        primary_data_type = self.get_primary_data_type()
        primary_data = self.get_data_type_data(primary_data_type)
        
        if not primary_data:
            print(f"No data available for {primary_data_type}")
            return
        
        # Process each item in the primary data type
        for item_key, item_data in primary_data.items():
            if not item_key:
                continue
            
            # Apply filters if they exist
            if hasattr(self, 'tag_filters') and self.tag_filters:
                continue_flag = False
                for header, filter_key in self.tag_filters:
                    acceptable_values = [value.strip() for value in filter_key.split(',')]
                    if item_data.get(header) not in acceptable_values:
                        continue_flag = True
                        break
                
                if continue_flag:
                    continue
            
            # Process coordinates for all data types
            data = {}
            
            # Process coordinates for each data type
            for data_type in self.get_all_data_types():
                coordinate_values = self.get_data_type_coordinate_values(data_type)
                data_type_data = self.get_data_type_data(data_type)
                
                if not data_type_data:
                    continue
                
                # For primary data type, use the item directly
                if data_type == primary_data_type:
                    for coordinate, value in coordinate_values.items():
                        raw_value = item_data.get(value)
                        data[coordinate] = raw_value
                else:
                    # For other data types, try to find matching data
                    # This is a simplified approach - you might need more complex logic
                    for coordinate, value in coordinate_values.items():
                        # Try to find matching data based on some key
                        # This is where you'd implement your specific logic
                        data[coordinate] = None  # Placeholder
            
            self.tag_cell_values[item_key] = data
        
        print("Coordinate Values generated (Dynamic):", self.tag_cell_values)

    def check_halt_flag(self):
        """Callback function to check if the process should be halted"""
        return self.halt_flag
    
    def reset_halt_flag(self):
        """Reset the halt flag to False"""
        self.halt_flag = False
    
    def set_halt_flag(self):
        """Set the halt flag to True to stop the process"""
        self.halt_flag = True
        print("Halt flag set - process will stop at next opportunity")
        if hasattr(self, 'status_label'):
            self.status_label.config(text="Stopping...", fg="orange")
    
    def update_status(self, message, color="black"):
        """Update the status label if it exists"""
        if hasattr(self, 'status_label'):
            self.status_label.config(text=message, fg=color)
    
    def add_datasheets(self):
        print('assigning tag coordinates')
        self.update_status("Assigning tag coordinates...", "blue")
        self.assign_value_coordinate_to_tag()
        print("Adding/Updating Datasheets")
        self.update_status("Adding/Updating Datasheets...", "blue")
        
        # Reset halt flag at the start
        self.reset_halt_flag()
        self.is_processing = True
        
        # Enable stop button and disable generate button
        if hasattr(self, 'stop_button'):
            self.stop_button.config(state="normal")
        if hasattr(self, 'generate_button'):
            self.generate_button.config(state="disabled")
        
        try:
            # Force reinitialization of Excel connection
            self.init_excel()
            print('Excel initialized')
            # Additional check for valid workbook reference
            if not self.excel_mgr.wb:
                raise Exception("Excel workbook not properly initialized")

            # Get the top tag for the primary data type
            primary_data_type = self.get_primary_data_type()
            primary_top_tag = self.get_data_type_top_tag(primary_data_type)
            print(f'primary_data_type: {primary_data_type}, top_tag: {primary_top_tag}')
            
            # Get the source sheet name for the primary data type
            primary_source_sheet_name = self.get_data_type_source_sheet_name(primary_data_type)
            
            # lets print all the variables that go into add_update_datasheets
            print('source_sheet_name', primary_source_sheet_name)
            print('tag_cell_values', self.tag_cell_values)
            print('datasheet_coord', self.datasheet_coord)
            print('ds_str', self.ds_str)
            print('rows_per_sheet', self.rows_per_sheet)

            # Check for potential naming conflict
            if primary_source_sheet_name and primary_source_sheet_name.startswith(self.ds_str):
                print(f"WARNING: Source sheet name '{primary_source_sheet_name}' starts with datasheet prefix '{self.ds_str}'")
                print("This may cause issues with sheet identification. Consider using a different prefix.")

            # Convert dropdown selection to the correct parameter value
            color_option = self.color_coding_var.get()
            if color_option == "None (Black)":
                color_option = None
            
            self.new_sheets = add_update_datasheets(self.excel_mgr.wb, primary_source_sheet_name,
                                            self.tag_cell_values, self.datasheet_coord,
                                            self.ds_str, rows_per_sheet=self.rows_per_sheet,
                                            key_coordinate=primary_top_tag,
                                            sig_figs=self.sig_figs, # Pass sig_figs
                                            tolerance=self.rounding_tolerance, # Pass tolerance
                                            halt_callback=self.check_halt_flag, # Pass halt callback
                                            cell_update_option=color_option) # Pass color coding option
            self.excel_mgr.mark_as_modified()
            
            if self.halt_flag:
                print("Process was halted by user")
                self.update_status("Process halted by user", "orange")
                messagebox.showinfo("Process Halted", "The datasheet addition process was halted by the user.")
            else:
                print("DONE")
                self.update_status("Process completed successfully", "green")
                
                # Generate and show detailed report
                self.show_generation_report()
        except Exception as e:
            print(f"Excel connection error: {e}")
            self.update_status("Error occurred", "red")
            messagebox.showerror("Excel Connection Error", 
                               "Excel connection lost. Please ensure Excel is open and try again.")
            # Reset Excel connection
            self.excel_mgr.wb = None
            self.excel_mgr.app = None
        finally:
            self.is_processing = False
            
            # Disable stop button and enable generate button
            if hasattr(self, 'stop_button'):
                self.stop_button.config(state="disabled")
            if hasattr(self, 'generate_button'):
                self.generate_button.config(state="normal")

    # endregion

    # region Data Loading and Saving

    def load_settings(self):
        """Load settings from a JSON file"""
        try:
            # Ask user to select a settings file
            file_path = filedialog.askopenfilename(
                title="Load Settings",
                filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
                initialdir="."
            )
            
            if not file_path:
                return
                
            with open(file_path, 'r', encoding='utf-8') as f:
                settings_data = json.load(f)
            
            # Validate the settings file format
            if 'version' not in settings_data:
                messagebox.showerror("Error", "Invalid settings file format.")
                return
            
            # Load entry values
            if 'entry_values' in settings_data:
                self.set_entry_values(settings_data['entry_values'])
            
            # Load default settings
            if 'default_settings' in settings_data:
                for key, value in settings_data['default_settings'].items():
                    if key in self.default_settings:
                        self.default_settings[key] = value
                        # Also set as instance attribute if it exists
                        if hasattr(self, key):
                            setattr(self, key, value)
            
            # Load dynamic attributes
            if 'dynamic_attributes' in settings_data:
                self.set_dynamic_attributes(settings_data['dynamic_attributes'])
            
            # Load data types configuration
            if 'data_types' in settings_data:
                for data_type, config in settings_data['data_types'].items():
                    # Create the data type if it doesn't exist
                    if data_type not in self.data_types:
                        self.add_data_type(data_type, config)
                    else:
                        # Update the existing data type configuration
                        for key, value in config.items():
                            if key in self.data_types[data_type]:
                                self.data_types[data_type][key] = value
            
            # Refresh the UI to reflect loaded settings
            # Clear the entries list since widgets will be recreated
            self.entries = []
            # First refresh the data sources notebook to show any new data types
            self.refresh_data_sources_notebook()
            # Then refresh data type frames in case new data types were loaded
            self.refresh_data_type_frames()
            # Finally refresh all tab content to show the loaded data
            self.refresh_tab_content(force_rebuild=True)
            
            messagebox.showinfo("Success", f"Settings loaded successfully from:\n{file_path}")
            
        except FileNotFoundError:
            messagebox.showerror("Error", "Settings file not found.")
        except json.JSONDecodeError as e:
            messagebox.showerror("Error", f"Invalid JSON file: {str(e)}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load settings: {str(e)}")
            print(f"Detailed error: {e}")
            import traceback
            traceback.print_exc()

    def save_settings(self, use_pickle=True):
        """Save current settings to a JSON file"""
        try:
            # Ask user where to save the settings file
            file_path = filedialog.asksaveasfilename(
                title="Save Settings",
                defaultextension=".json",
                filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
                initialfile=f"datasheet_helper_settings_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
            )
            
            if not file_path:
                return
            
            # Collect all settings data
            settings_data = {
                'version': '1.0',
                'saved_date': datetime.now().isoformat(),
                'entry_values': self.get_entry_values(),
                'default_settings': self.default_settings.copy(),
                'dynamic_attributes': self.get_dynamic_attributes(),
                'data_types': {}
            }
            
            # Save data types configuration
            for data_type, config in self.data_types.items():
                settings_data['data_types'][data_type] = {
                    'headers': config.get('headers', []),
                    'coordinate_values': config.get('coordinate_values', {}),
                    'selected_sheets': config.get('selected_sheets', None),
                    'top_tag': config.get('top_tag', 'A1'),
                    'source_sheet_name': config.get('source_sheet_name', 'TEMPLATE'),
                    'data': config.get('data', {}),
                    'path_key': config.get('path_key', f'{data_type}_path')
                }
            
            # Write to file
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(settings_data, f, indent=2, ensure_ascii=False)
            
            messagebox.showinfo("Success", f"Settings saved successfully to:\n{file_path}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save settings: {str(e)}")
            print(f"Detailed error: {e}")
            import traceback
            traceback.print_exc()

    def get_entry_values(self):
        """Get current values from all entry widgets"""
        entry_values = {}
        
        # Get values from self.entries (main entry widgets)
        for entry, variable in self.entries:
            try:
                # Handle different widget types
                if hasattr(entry, 'get'):
                    entry_values[variable] = entry.get()
                elif hasattr(entry, 'selection_get'):
                    # Handle text widgets
                    entry_values[variable] = entry.selection_get()
                else:
                    entry_values[variable] = str(entry)
            except Exception as e:
                print(f"Error getting entry value for {variable}: {e}")
                entry_values[variable] = ""
        
        return entry_values
    
    def set_entry_values(self, entry_values):
        """Set values to all entry widgets"""
        for variable, value in entry_values.items():
            try:
                # Find the corresponding entry widget
                for entry, var in self.entries:
                    if var == variable:
                        # Handle different widget types
                        if hasattr(entry, 'delete') and hasattr(entry, 'insert'):
                            # Standard Entry widget
                            entry.delete(0, tk.END)
                            entry.insert(0, str(value))
                        elif hasattr(entry, 'set'):
                            # Combobox widget
                            entry.set(str(value))
                        elif hasattr(entry, 'delete') and hasattr(entry, 'insert'):
                            # Text widget
                            entry.delete(1.0, tk.END)
                            entry.insert(1.0, str(value))
                        break
            except Exception as e:
                print(f"Error setting entry value for {variable}: {e}")
    
    def get_dynamic_attributes(self):
        """Get any additional dynamic attributes that aren't in default_settings"""
        dynamic_attrs = {}
        
        # Get attributes that might not be in default_settings but are important
        important_attrs = [
            'current_transform_data_type', 'current_transform_key',
            'tag_filters', 'tag_cell_values', 'coordinate_conversions', 
            'coordinate_combinations', 'semantic_model', 'model_loaded', 'loading_model'
        ]
        
        for attr in important_attrs:
            if hasattr(self, attr):
                try:
                    value = getattr(self, attr)
                    # Only save if it's serializable
                    if self._is_serializable(value):
                        dynamic_attrs[attr] = value
                except Exception as e:
                    print(f"Error getting dynamic attribute {attr}: {e}")
        
        return dynamic_attrs
    
    def set_dynamic_attributes(self, dynamic_attrs):
        """Set dynamic attributes from saved data"""
        for attr, value in dynamic_attrs.items():
            try:
                setattr(self, attr, value)
            except Exception as e:
                print(f"Error setting dynamic attribute {attr}: {e}")
    
    def _is_serializable(self, obj):
        """Check if an object is JSON serializable"""
        try:
            json.dumps(obj)
            return True
        except (TypeError, ValueError):
            return False

    
    def load_data_type_from_datasheet(self, data_type, file_path=None):
        """Load data for any data type from datasheet"""
        config = self.get_data_type_config(data_type)
        name = data_type
        
        def set_data(data):
            self.set_data_type_data(data_type, data)
            self.refresh_tab_content()
            print(f"{name} data loaded from datasheet")
        
        app_window = tk.Toplevel(self.root)
        DatasheetExtractor(app_window, callback=set_data, file_path=file_path)

    def load_data_type_from_json(self, data_type):
        """Load data for any data type from JSON file"""
        config = self.get_data_type_config(data_type)
        name = data_type
        
        # Ask the user to select a JSON file
        file_path = filedialog.askopenfilename(
            title=f"Select JSON file for {name}",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        
        if file_path:
            data = load_dict_from_json(file_path)
            self.set_data_type_data(data_type, data)
            self.refresh_tab_content()
            print(f"{name} data loaded from JSON")

    def update_data_type_keys(self, data_type):
        """Update keys for any data type using transformation code"""
        config = self.get_data_type_config(data_type)
        name = data_type
        
        code = askstring(f"Enter transformation code for keys in {name}", 
                        f"Enter transformation code for {name}",
                        initialvalue='"-".join(x.split("-")[-2:])')
        
        if code:
            data = self.get_data_type_data(data_type)
            transformed_data = transform_dictionary(data, code)
            self.set_data_type_data(data_type, transformed_data)
            print(f"{name} keys updated")

    # endregion

    # region Data Display

    def view_data(self, text):

        print(f"Viewing {text}")

        if text == "Coordinate-Value Data":
            # open a new tkinter popup window with a scroll bar showing all the coordinate-value pairs
            self.display_coordinate_values()
        elif text == "Datasheets":
            # open the datasheets file
            os.startfile(self.datasheets)
        else:
            # Check if it's a data type
            data_type = self.get_data_type_by_name(text)
            if data_type:
                self.display_data_type(data_type)
            else:
                print(f"Unknown data type: {text}")
    
    def display_data_type(self, data_type):
        """Display data for any data type using the centralized system"""
        config = self.get_data_type_config(data_type)
        name = data_type
        data = self.get_data_type_data(data_type)
        
        # Create a new window
        view_window = tk.Toplevel(self.root)
        view_window.title(name)

        # Create a button frame for actions
        button_frame = tk.Frame(view_window)
        button_frame.pack(fill=tk.X, padx=10, pady=5)

        def refresh_display():
            """Refresh the display with current data"""
            scrolled_text.configure(state='normal')
            scrolled_text.delete(1.0, tk.END)
            
            current_data = self.get_data_type_data(data_type)
            if current_data:
                for key, value in current_data.items():
                    scrolled_text.insert(tk.END, f"{key}: {value}\n\n")
            else:
                scrolled_text.insert(tk.END, f"No {name} data available.")
            
            scrolled_text.configure(state='disabled')

        def split_keys():
            """Split composite keys based on delimiter"""
            # Ask for delimiter
            delimiter = askstring("Split Keys", 
                                "Enter delimiter to split keys on (e.g., ';', ',', '|'):",
                                initialvalue=";")
            
            if not delimiter:
                return
            
            # Get current data
            current_data = self.get_data_type_data(data_type)
            if not current_data:
                tk.messagebox.showwarning("No Data", f"No {name} data available to split.")
                return
            
            # Find keys that contain the delimiter
            keys_to_split = []
            for key in current_data.keys():
                if delimiter in str(key):
                    keys_to_split.append(key)
            
            if not keys_to_split:
                tk.messagebox.showinfo("No Split Needed", 
                                     f"No keys found containing the delimiter '{delimiter}'.")
                return
            
            # Confirm the split operation
            if len(keys_to_split) == 1:
                message = f"Found 1 key to split: '{keys_to_split[0]}'\n\nThis will create separate entries for each part."
            else:
                message = f"Found {len(keys_to_split)} keys to split:\n"
                for key in keys_to_split[:5]:  # Show first 5 keys
                    message += f"  - {key}\n"
                if len(keys_to_split) > 5:
                    message += f"  ... and {len(keys_to_split) - 5} more\n"
                message += "\nThis will create separate entries for each part."
            
            if not tk.messagebox.askyesno("Confirm Split", message):
                return
            
            # Perform the split
            new_data = {}
            split_count = 0
            
            for key, value in current_data.items():
                if delimiter in str(key):
                    # Split the key
                    key_parts = str(key).split(delimiter)
                    for part in key_parts:
                        part = part.strip()  # Remove leading/trailing whitespace
                        if part:  # Only add non-empty parts
                            new_data[part] = value  # Use the same value for all split keys
                            split_count += 1
                else:
                    # Keep the original key unchanged
                    new_data[key] = value
            
            # Update the data
            self.set_data_type_data(data_type, new_data)
            
            # Show success message
            tk.messagebox.showinfo("Split Complete", 
                                 f"Successfully split {len(keys_to_split)} keys into {split_count} new entries.")
            
            # Refresh the display
            refresh_display()

        def paste_from_clipboard():
            """Paste JSON data from clipboard and overwrite current data"""
            try:
                # Get clipboard content
                clipboard_content = self.root.clipboard_get()
                
                # Try to parse as JSON
                pasted_data = json.loads(clipboard_content)
                
                # Validate that it's a dictionary
                if not isinstance(pasted_data, dict):
                    tk.messagebox.showerror("Invalid Format", 
                                          "Clipboard content must be a JSON object (dictionary).")
                    return
                
                # Confirm overwrite
                result = tk.messagebox.askyesno("Confirm Overwrite", 
                                              f"This will overwrite all current {name} data with the clipboard content.\n\n"
                                              f"Found {len(pasted_data)} entries in clipboard.\n\n"
                                              "Do you want to continue?")
                
                if result:
                    # Update the data
                    self.set_data_type_data(data_type, pasted_data)
                    
                    # Show success message
                    tk.messagebox.showinfo("Paste Complete", 
                                         f"Successfully pasted {len(pasted_data)} entries from clipboard.")
                    
                    # Refresh the display
                    refresh_display()
                    
            except tk.TclError:
                tk.messagebox.showerror("Clipboard Error", 
                                      "No content found in clipboard.")
            except json.JSONDecodeError as e:
                tk.messagebox.showerror("Invalid JSON", 
                                      f"Clipboard content is not valid JSON:\n{str(e)}")
            except Exception as e:
                tk.messagebox.showerror("Error", 
                                      f"An error occurred while pasting from clipboard:\n{str(e)}")

        # Add Split Keys button
        split_button = tk.Button(button_frame, text="Split Keys", command=split_keys)
        split_button.pack(side=tk.LEFT, padx=5)
        
        # Add Paste from Clipboard button
        paste_button = tk.Button(button_frame, text="Paste from Clipboard", command=paste_from_clipboard)
        paste_button.pack(side=tk.LEFT, padx=5)
        
        # Add Modify Keys button
        modify_keys_button = tk.Button(button_frame, text="Modify Keys", command=lambda: self.update_data_type_keys(data_type))
        modify_keys_button.pack(side=tk.LEFT, padx=5)

        # Create a scrolled text widget to display the data
        scrolled_text = scrolledtext.ScrolledText(view_window, width=40, height=20)
        scrolled_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)  # Fill and expand to fill the window

        # Display the data content in the scrolled text widget
        if data:
            for key, value in data.items():
                scrolled_text.insert(tk.END, f"{key}: {value}\n\n")
        else:
            scrolled_text.insert(tk.END, f"No {name} data available.")

        scrolled_text.configure(state='disabled')  # Make read-only



    def display_coordinate_values(self):

        # Create a new window
        view_window = tk.Toplevel(self.root)
        view_window.title("Coordinate values")

        coord_label = tk.Label(view_window, text='Index Data Coordinates')
        scrolled_text1 = scrolledtext.ScrolledText(view_window, width=40, height=20)
        coord_label.pack()
        scrolled_text1.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)  # Fill and expand to fill the window

        # Display the coordinate content in the scrolled text widget
        if self.tag_cell_values:
            for key, value in self.tag_cell_values.items():
                scrolled_text1.insert(tk.END, f"{key}: {value}\n\n")
        else:
            scrolled_text1.insert(tk.END, "No coordinate data available.")

        scrolled_text1.configure(state='disabled')  # Make

    def display_coordinate_values(self):

        # Create a new window
        view_window = tk.Toplevel(self.root)
        view_window.title("Coordinate values")

        coord_label = tk.Label(view_window, text='Index Data Coordinates')
        scrolled_text1 = scrolledtext.ScrolledText(view_window, width=40, height=20)
        coord_label.pack()
        scrolled_text1.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)  # Fill and expand to fill the window

        # Display the coordinate content in the scrolled text widget
        if self.tag_cell_values:
            for key, value in self.tag_cell_values.items():
                scrolled_text1.insert(tk.END, f"{key}: {value}\n\n")
        else:
            scrolled_text1.insert(tk.END, "No coordinate data available.")

        scrolled_text1.configure(state='disabled')  # Make

    # endregion

    # region Excel Operations

    def save_and_close_workbook(self):
        """Save and close the current workbook"""
        self.excel_mgr.close_workbook()

    def release_excel_connection(self):
        """Release the xlwings connection to Excel, allowing user to save independently"""
        try:
            if self.excel_mgr.wb or self.excel_mgr.app:
                # Release the connection but keep Excel open
                self.excel_mgr.release_connection()
                tk.messagebox.showinfo("Excel Released", 
                                     "xlwings connection has been released. The workbook remains open in Excel and you can now save it independently.")
            else:
                tk.messagebox.showinfo("No Connection", 
                                     "No Excel connection is currently active.")
        except Exception as e:
            tk.messagebox.showerror("Error", 
                                  f"Error releasing Excel connection: {str(e)}")

    def show_generation_report(self):
        """Show a detailed report of the datasheet generation process"""
        try:
            # Count the number of sheets created/updated
            num_sheets = len(self.new_sheets) if self.new_sheets else 0
            
            # Get additional information for the report
            source_sheet = self.source_sheet_name if hasattr(self, 'source_sheet_name') else "Unknown"
            datasheet_prefix = self.ds_str if hasattr(self, 'ds_str') else "Unknown"
            rows_per_sheet = self.rows_per_sheet if hasattr(self, 'rows_per_sheet') else "Unknown"
            
            # Create the report message
            report_message = f"""Datasheet Generation Complete!

📊 Generation Summary:
• Source Sheet: {source_sheet}
• Datasheet Prefix: {datasheet_prefix}
• Sheets Created/Updated: {num_sheets}
• Rows per Sheet: {rows_per_sheet}

✅ Process completed successfully!
The datasheets have been generated and are ready for use."""
            
            # Show the report in a message box
            tk.messagebox.showinfo("Generation Report", report_message)
            
        except Exception as e:
            # Fallback to a simple success message if there's an error
            tk.messagebox.showinfo("Generation Complete", 
                                 "Datasheet generation completed successfully!")
            print(f"Error creating detailed report: {e}")

    def delete_added_sheets(self):
        if self.excel_mgr.wb:
            for sheet in self.new_sheets:
                self.excel_mgr.wb.sheets[sheet].delete()
            self.excel_mgr.mark_as_modified()

    def open_excel_search_app(self):
        # Instantiate and show the ExcelSearchApp
        excel_search_app = ExcelSearchApp()
        excel_search_app.mainloop()

    def open_edit_xlsx(self):
        edit_xlsx_window = tk.Toplevel(self.root)
        ExcelEditorApp(edit_xlsx_window)

    def get_sheet_names(self, force_reload=False):
        print("DEBUG: Starting get_sheet_names...")
        if not self.datasheets:
            print("DEBUG: No datasheets file set, returning empty list")
            return []
        
        # Check if we have cached sheet names and they're still valid
        if (not force_reload and 
            hasattr(self, '_cached_sheet_names') and 
            hasattr(self, '_cached_sheet_file') and 
            self._cached_sheet_file == self.datasheets):
            print("DEBUG: Using cached sheet names")
            return self._cached_sheet_names
        
        try:
            print(f"DEBUG: Loading workbook: {self.datasheets}")
            # Use read_only=True and data_only=True for faster loading
            wb = openpyxl.load_workbook(self.datasheets, read_only=True, data_only=True)
            print("DEBUG: Workbook loaded successfully")
            sheet_names = wb.sheetnames
            print(f"DEBUG: Found {len(sheet_names)} sheets: {sheet_names}")
            wb.close()  # Explicitly close the workbook
            print("DEBUG: Workbook closed")
            
            # Cache the sheet names
            self._cached_sheet_names = sheet_names
            self._cached_sheet_file = self.datasheets
            print("DEBUG: Sheet names cached")
            return sheet_names
        except Exception as e:
            print(f"DEBUG: Error getting sheet names from {self.datasheets}: {e}")
            return []

    def on_closing(self):
        # Stop Excel monitoring
        self.stop_excel_monitoring()
        
        # Release Excel connection before cleanup
        try:
            self.release_excel_connection()
        except Exception as e:
            print(f"Warning: Error releasing Excel connection: {e}")
        
        self.root.destroy()


    # endregion

    # region UI Utilities

    def update_entry(self, entry, variable):
        filename = ''
        if variable == "coordinate_value":
            filename = self.coordinate_value_path
        elif variable == "datasheets":
            filename = self.datasheets
        else:
            # Check if it's a data type path
            for data_type in self.get_all_data_types():
                config = self.get_data_type_config(data_type)
                path_key = config.get('path_key', f'{data_type}_path')
                if variable == path_key:
                    filename = getattr(self, path_key, '')
                    break

        entry.delete(0, tk.END)
        entry.insert(0, filename)

    def update_entries(self):
        for entry, entry_var in self.entries:
            try:
                self.update_entry(entry, entry_var)
                entry.xview_moveto(1)
            except Exception as e:
                print(f'error {e}')

    def browse_destination(self, entry):
        """Browse for destination datasheet file"""
        print("DEBUG: Starting browse_destination...")
        filename = filedialog.askopenfilename(
            filetypes=[("All supported files", "*.json;*.xlsx;*.xls"), 
                      ("JSON files", "*.json"), 
                      ("Excel files", "*.xlsx;*.xls"),
                      ("All files", "*.*")]
        )
        
        # Only update the entry if a file was actually selected (not canceled)
        if filename:
            try:
                print("DEBUG: File selected, updating entry...")
                entry.delete(0, tk.END)
                entry.insert(0, filename)
                self.datasheets = filename
                print(f"DEBUG: Destination set to: {self.datasheets}")
                
                # Clear cached sheet names since we have a new file
                if hasattr(self, '_cached_sheet_names'):
                    self._cached_sheet_names = None
                    self._cached_sheet_file = None
                    print("DEBUG: Cleared cached sheet names for new file")
                
                # Test if the file can be opened without freezing
                print("DEBUG: Testing file accessibility...")
                test_sheets = self.get_sheet_names(force_reload=True)
                if test_sheets:
                    print(f"DEBUG: File accessible, found {len(test_sheets)} sheets")
                else:
                    print("DEBUG: Warning: Could not read sheet names from file")
                
                print("DEBUG: File accessibility test completed")
                    
            except Exception as e:
                print(f"DEBUG: Error setting destination file: {e}")
                # Still set the filename even if there's an error
                self.datasheets = filename
        
        print("DEBUG: browse_destination completed")

    def browse(self, entry, variable):
        # Allow both JSON and Excel files
        filename = filedialog.askopenfilename(
            filetypes=[("All supported files", "*.json;*.xlsx;*.xls"), 
                      ("JSON files", "*.json"), 
                      ("Excel files", "*.xlsx;*.xls"),
                      ("All files", "*.*")]
        )
        
        # Only update the entry if a file was actually selected (not canceled)
        if filename:
            entry.delete(0, tk.END)
            entry.insert(0, filename)

        # Only update instance variables if a file was actually selected
        if filename:
            if variable == "coordinate_value":
                self.coordinate_value_path = filename
            elif variable == "datasheets":
                self.datasheets = filename
                print(self.datasheets)
            else:
                # Check if it's a data type path
                for data_type in self.get_all_data_types():
                    config = self.get_data_type_config(data_type)
                    path_key = config.get('path_key', f'{data_type}_path')
                    if variable == path_key:
                        setattr(self, path_key, filename)
                        print(f"Set {path_key} to {filename}")
                        break

    def open_file(self, entry):
        """Open the file specified in the entry using the system default application"""
        file_path = entry.get().strip()
        if not file_path:
            messagebox.showwarning("No File", "Please specify a file path first.")
            return
        
        if not os.path.exists(file_path):
            messagebox.showerror("File Not Found", f"The file '{file_path}' does not exist.")
            return
        
        try:
            os.startfile(file_path)
        except Exception as e:
            messagebox.showerror("Error Opening File", f"Could not open file: {str(e)}")

    def configure(self, text):
        print(f"Configure {text}")

            # Find the data type by name
        data_type = self.get_data_type_by_name(text)
        if data_type:
            self.configure_data_type(data_type)
        else:
            print(f"Unknown configuration type: {text}")
    
    def get_data_type_by_name(self, name):
        """Get data type key by its display name"""
        # Case-sensitive matching
        for data_type in self.data_types.keys():
            if data_type == name:
                return data_type
        return None
    
    def configure_data_type(self, data_type):
        """Configure a specific data type using the centralized system"""
        config = self.get_data_type_config(data_type)
        name = data_type
        path_key = config.get('path_key', f'{data_type}_path')
        
        file_path = getattr(self, path_key, '')
        current_headers = self.get_data_type_headers(data_type)
        current_selection = self.get_data_type_selected_sheets(data_type)
        current_tolerance = self.blank_cell_tolerance

        if not file_path or not os.path.exists(file_path):
            messagebox.showwarning("File Not Found", f"Please select a valid {name} file first.", parent=self.root)
            return

        result = configure_source_data_dialog(self.root, f"Configure {name} Source",
                                              current_headers, file_path, current_selection, current_tolerance)

        if result:
            self.set_data_type_headers(data_type, result["headers"])
            self.set_data_type_selected_sheets(data_type, result["selected_sheets"])
            self.blank_cell_tolerance = result["tolerance"]
            print(f"{name} configuration updated.")


    def configure_ds(self):
        print("DEBUG: Starting configure_ds...")
        self.init_excel()
        print("DEBUG: init_excel completed, starting refresh_tab_content...")
        self.refresh_tab_content(force_rebuild=True)
        print("DEBUG: refresh_tab_content completed, starting update_entries...")
        self.update_entries()
        print("DEBUG: configure_ds completed")

    def delete_sheets_by_prefix(self):
        """Delete all sheets that start with a user-defined prefix"""
        if not self.excel_mgr.wb:
            print("No workbook is currently open")
            return

        prefix = askstring("Delete Sheets", "Enter the prefix of sheets to delete:", initialvalue="DS-")
        if not prefix:
            return

        sheets_to_delete = []
        for sheet in self.excel_mgr.wb.sheets:
            if sheet.name.startswith(prefix):
                sheets_to_delete.append(sheet.name)

        if not sheets_to_delete:
            print(f"No sheets found starting with prefix '{prefix}'")
            return

        if messagebox.askyesno("Confirm Delete",
                               f"Are you sure you want to delete {len(sheets_to_delete)} sheets starting with '{prefix}'?"):
            for sheet_name in sheets_to_delete:
                self.excel_mgr.wb.sheets[sheet_name].delete()
            self.excel_mgr.mark_as_modified()
            print(f"Deleted {len(sheets_to_delete)} sheets")

    # endregion

    # region Unit Conversion

    def open_conversion_dialog(self):
        if not self.selected_coord_for_context:
            return

        coord = self.selected_coord_for_context
        existing_conversion = self.coordinate_conversions.get(coord, {})

        dialog = Toplevel(self.root)
        dialog.title(f"Unit Conversion for {coord}")
        
        # Get cursor position and set dialog position near cursor
        cursor_x = self.root.winfo_pointerx()
        cursor_y = self.root.winfo_pointery()
        
        # Add some offset so dialog doesn't appear exactly at cursor
        dialog_x = cursor_x + 10
        dialog_y = cursor_y + 10
        
        # Ensure dialog doesn't go off screen
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        dialog_width = 350
        dialog_height = 250
        
        if dialog_x + dialog_width > screen_width:
            dialog_x = screen_width - dialog_width - 10
        if dialog_y + dialog_height > screen_height:
            dialog_y = screen_height - dialog_height - 10
            
        dialog.geometry(f"{dialog_width}x{dialog_height}+{dialog_x}+{dialog_y}")
        dialog.transient(self.root)
        dialog.grab_set()

        # Predefined common units
        common_units = ["", "deg F", "deg C", "psi", "bar", "kPa", "in", "mm", "ft", "m", "gal", "L"]
        # Predefined common formulas (string representation for storage)
        predefined_formulas = {
            "deg F -> deg C": "lambda x: (x - 32) * 5 / 9",
            "deg C -> deg F": "lambda x: (x * 9 / 5) + 32",
            "psi -> bar": "lambda x: x / 14.5038",
            "bar -> psi": "lambda x: x * 14.5038",
            "psi -> kPa": "lambda x: x * 6.89476",
            "kPa -> psi": "lambda x: x / 6.89476",
            "in -> mm": "lambda x: x * 25.4",
            "mm -> in": "lambda x: x / 25.4",
            "ft -> m": "lambda x: x * 0.3048",
            "m -> ft": "lambda x: x / 0.3048",
            "gal -> L": "lambda x: x * 3.78541",
            "L -> gal": "lambda x: x / 3.78541",
            "lb/hr -> kg/hr": "lambda x: x * 2.20462",
            "kg/hr -> lb/hr": "lambda x: x / 2.20462",
            "scfm -> m3/hr": "lambda x: x * 0.0283168",
            "m3/hr -> scfm": "lambda x: x / 0.0283168"
        }

        main_frame = ttk.Frame(dialog, padding="10")
        main_frame.pack(fill=BOTH, expand=True)

        # Input Unit
        ttk.Label(main_frame, text="Input Unit:").grid(row=0, column=0, sticky=W, pady=2)
        in_unit_combo = ttk.Combobox(main_frame, values=common_units)
        in_unit_combo.grid(row=0, column=1, sticky=EW, padx=5, pady=2)
        in_unit_combo.set(existing_conversion.get('in_unit', ''))

        # Output Unit
        ttk.Label(main_frame, text="Output Unit:").grid(row=1, column=0, sticky=W, pady=2)
        out_unit_combo = ttk.Combobox(main_frame, values=common_units)
        out_unit_combo.grid(row=1, column=1, sticky=EW, padx=5, pady=2)
        out_unit_combo.set(existing_conversion.get('out_unit', ''))

        # Formula Selection
        ttk.Label(main_frame, text="Formula:").grid(row=2, column=0, sticky=NW, pady=2)
        formula_frame = ttk.Frame(main_frame)
        formula_frame.grid(row=2, column=1, sticky=EW, padx=5, pady=2)

        formula_var = tk.StringVar(value=existing_conversion.get('formula', ''))
        formula_entry = ttk.Entry(formula_frame, textvariable=formula_var, width=30)
        formula_entry.pack(side=LEFT, fill=X, expand=True)

        # Predefined Formula Dropdown
        predefined_combo = ttk.Combobox(main_frame, values=list(predefined_formulas.keys()), state="readonly")
        predefined_combo.grid(row=3, column=1, sticky=EW, padx=5, pady=5)
        predefined_combo.set("Select Predefined...") # Placeholder

        def apply_predefined(event):
            selected_key = predefined_combo.get()
            if selected_key in predefined_formulas:
                formula_var.set(predefined_formulas[selected_key])
                # Attempt to set units based on common pattern "unit1 -> unit2"
                try:
                    units = selected_key.split(' -> ')
                    in_unit_combo.set(units[0])
                    out_unit_combo.set(units[1])
                except:
                    pass # Ignore errors if parsing fails

        predefined_combo.bind("<<ComboboxSelected>>", apply_predefined)

        # Buttons
        button_frame = ttk.Frame(dialog)
        button_frame.pack(side=BOTTOM, fill=X, padx=10, pady=10)

        def save_conversion():
            formula_str = formula_var.get().strip()
            if formula_str:
                # Basic validation: check if it starts with lambda x:
                if not formula_str.startswith('lambda x:'):
                    messagebox.showerror("Invalid Formula", "Formula must start with 'lambda x:'", parent=dialog)
                    return
                # Attempt a dummy eval to catch syntax errors
                try:
                    eval(formula_str)(1) # Test with a dummy value
                except Exception as e:
                    messagebox.showerror("Invalid Formula", f"Formula error: {e}", parent=dialog)
                    return

                self.coordinate_conversions[coord] = {
                    'in_unit': in_unit_combo.get(),
                    'out_unit': out_unit_combo.get(),
                    'formula': formula_str
                }
                print(f"Saved conversion for {coord}: {self.coordinate_conversions[coord]}")
            else:
                # If formula is empty, remove the conversion
                if coord in self.coordinate_conversions:
                    del self.coordinate_conversions[coord]
                    print(f"Removed conversion for {coord} (empty formula)")

            # Update listboxes in the main window (call the method on the tab)
            coordinates_tab = self.notebook.tabs()[self.notebook.index('current')] # Assuming Coordinates is current, maybe fragile
            # A more robust way might be needed if tab order changes
            for i, tab_name in enumerate(self.notebook.tabs()):
                if self.notebook.tab(i, "text") == 'Coordinates':
                    coordinates_tab = self.notebook.winfo_children()[i]
                    # A more robust way might be needed if tab order changes
                    if hasattr(coordinates_tab, 'reinitialize'): # Check if reinitialize exists
                         coordinates_tab.reinitialize() # Refresh listboxes
                    break

            dialog.destroy()

        ttk.Button(button_frame, text="OK", command=save_conversion).pack(side=RIGHT, padx=5)
        ttk.Button(button_frame, text="Cancel", command=dialog.destroy).pack(side=RIGHT)

        main_frame.columnconfigure(1, weight=1)
        self.root.wait_window(dialog)

    def remove_conversion(self):
        if self.selected_coord_for_context and self.selected_coord_for_context in self.coordinate_conversions:
            coord = self.selected_coord_for_context
            del self.coordinate_conversions[coord]
            print(f"Removed conversion for {coord}")
            # Update listboxes
            for i, tab_name in enumerate(self.notebook.tabs()):
                if self.notebook.tab(i, "text") == 'Coordinates':
                    coordinates_tab = self.notebook.winfo_children()[i]
                    if hasattr(coordinates_tab, 'reinitialize'):
                         coordinates_tab.reinitialize()
                    break
        else:
            print("No conversion selected or found to remove.")

    def open_combination_dialog(self):
        if not self.selected_coord_for_context:
            return

        coord = self.selected_coord_for_context
        existing_combination = self.coordinate_combinations.get(coord, {})

        dialog = Toplevel(self.root)
        dialog.title(f"Coordinate Combination for {coord}")
        
        # Get cursor position and set dialog position near cursor
        cursor_x = self.root.winfo_pointerx()
        cursor_y = self.root.winfo_pointery()
        
        # Add some offset so dialog doesn't appear exactly at cursor
        dialog_x = cursor_x + 10
        dialog_y = cursor_y + 10
        
        # Ensure dialog doesn't go off screen
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        dialog_width = 400
        dialog_height = 300
        
        if dialog_x + dialog_width > screen_width:
            dialog_x = screen_width - dialog_width - 10
        if dialog_y + dialog_height > screen_height:
            dialog_y = screen_height - dialog_height - 10
            
        dialog.geometry(f"{dialog_width}x{dialog_height}+{dialog_x}+{dialog_y}")
        dialog.transient(self.root)
        dialog.grab_set()

        main_frame = ttk.Frame(dialog, padding="10")
        main_frame.pack(fill=BOTH, expand=True)

        # Get all available coordinates for selection
        all_coordinates = []
        for data_type in self.get_all_data_types():
            coord_values = self.get_data_type_coordinate_values(data_type)
            all_coordinates.extend(list(coord_values.keys()))
        # Remove the current coordinate from the list to avoid self-reference
        if coord in all_coordinates:
            all_coordinates.remove(coord)

        # Operation selection
        ttk.Label(main_frame, text="Operation:").grid(row=0, column=0, sticky=W, pady=2)
        operation_combo = ttk.Combobox(main_frame, values=["add", "subtract", "multiply", "divide", "concatenate"], state="readonly")
        operation_combo.grid(row=0, column=1, sticky=EW, padx=5, pady=2)
        operation_combo.set(existing_combination.get('operation', 'add'))

        # Coordinates to combine
        ttk.Label(main_frame, text="Coordinates to combine:").grid(row=1, column=0, sticky=NW, pady=2)
        
        # Create a frame for the listbox and scrollbar
        list_frame = ttk.Frame(main_frame)
        list_frame.grid(row=1, column=1, sticky=EW, padx=5, pady=2)
        
        # Listbox for selected coordinates
        selected_listbox = tk.Listbox(list_frame, height=6, width=30)
        selected_listbox.pack(side=LEFT, fill=BOTH, expand=True)
        
        # Scrollbar for the listbox
        scrollbar = ttk.Scrollbar(list_frame, orient=VERTICAL, command=selected_listbox.yview)
        scrollbar.pack(side=RIGHT, fill=Y)
        selected_listbox.config(yscrollcommand=scrollbar.set)

        # Populate with existing combinations
        existing_combines = existing_combination.get('combines', [])
        for coord_name in existing_combines:
            selected_listbox.insert(tk.END, coord_name)

        # Available coordinates dropdown
        ttk.Label(main_frame, text="Add coordinate:").grid(row=2, column=0, sticky=W, pady=2)
        coord_combo = ttk.Combobox(main_frame, values=all_coordinates, state="readonly")
        coord_combo.grid(row=2, column=1, sticky=EW, padx=5, pady=2)

        def add_coordinate():
            selected = coord_combo.get()
            if selected and selected not in [selected_listbox.get(i) for i in range(selected_listbox.size())]:
                selected_listbox.insert(tk.END, selected)
                coord_combo.set('')

        def remove_coordinate():
            selected_indices = selected_listbox.curselection()
            if selected_indices:
                # Remove in reverse order to avoid index shifting
                for index in sorted(selected_indices, reverse=True):
                    selected_listbox.delete(index)

        # Buttons for adding/removing coordinates
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=3, column=0, columnspan=2, pady=5)
        
        ttk.Button(button_frame, text="Add", command=add_coordinate).pack(side=LEFT, padx=2)
        ttk.Button(button_frame, text="Remove Selected", command=remove_coordinate).pack(side=LEFT, padx=2)

        # Main dialog buttons
        dialog_button_frame = ttk.Frame(dialog)
        dialog_button_frame.pack(side=BOTTOM, fill=X, padx=10, pady=10)

        def save_combination():
            operation = operation_combo.get()
            combines = [selected_listbox.get(i) for i in range(selected_listbox.size())]
            
            if not operation:
                messagebox.showerror("Invalid Input", "Please select an operation.", parent=dialog)
                return
                
            if not combines:
                messagebox.showerror("Invalid Input", "Please select at least one coordinate to combine.", parent=dialog)
                return

            self.coordinate_combinations[coord] = {
                'operation': operation,
                'combines': combines
            }
            print(f"Saved combination for {coord}: {self.coordinate_combinations[coord]}")

            # Update listboxes
            for i, tab_name in enumerate(self.notebook.tabs()):
                if self.notebook.tab(i, "text") == 'Coordinates':
                    coordinates_tab = self.notebook.winfo_children()[i]
                    if hasattr(coordinates_tab, 'reinitialize'):
                         coordinates_tab.reinitialize()
                    break

            dialog.destroy()

        ttk.Button(dialog_button_frame, text="OK", command=save_combination).pack(side=RIGHT, padx=5)
        ttk.Button(dialog_button_frame, text="Cancel", command=dialog.destroy).pack(side=RIGHT)

        main_frame.columnconfigure(1, weight=1)
        self.root.wait_window(dialog)

    def remove_combination(self):
        if self.selected_coord_for_context and self.selected_coord_for_context in self.coordinate_combinations:
            coord = self.selected_coord_for_context
            del self.coordinate_combinations[coord]
            print(f"Removed combination for {coord}")
            # Update listboxes
            for i, tab_name in enumerate(self.notebook.tabs()):
                if self.notebook.tab(i, "text") == 'Coordinates':
                    coordinates_tab = self.notebook.winfo_children()[i]
                    if hasattr(coordinates_tab, 'reinitialize'):
                         coordinates_tab.reinitialize()
                    break
        else:
            print("No combination selected or found to remove.")

    def show_excel_processing_dialog(self, data_type, file_path):
        """Show custom dialog for Excel file processing options"""
        config = self.get_data_type_config(data_type)
        name = data_type
        
        # Create dialog window
        dialog = tk.Toplevel(self.root)
        dialog.title("Excel File Processing")
        dialog.geometry("500x500")
        dialog.resizable(False, False)
        
        # Center the dialog
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Main frame
        main_frame = ttk.Frame(dialog, padding="20")
        main_frame.pack(fill="both", expand=True)
        
        # Title
        title_label = ttk.Label(main_frame, text=f"Process Excel File for {name}", 
                               font=("Arial", 12, "bold"))
        title_label.pack(pady=(0, 10))
        
        # File path display
        file_frame = ttk.Frame(main_frame)
        file_frame.pack(fill="x", pady=(0, 15))
        
        ttk.Label(file_frame, text="File:", font=("Arial", 9, "bold")).pack(anchor="w")
        file_label = ttk.Label(file_frame, text=file_path, foreground="blue", 
                              font=("Arial", 9))
        file_label.pack(anchor="w", pady=(2, 0))
        
        # Description
        desc_frame = ttk.Frame(main_frame)
        desc_frame.pack(fill="x", pady=(0, 20))
        
        ttk.Label(desc_frame, text="How would you like to process this Excel file?", 
                 font=("Arial", 10)).pack(anchor="w", pady=(0, 10))
        
        # Options frame
        options_frame = ttk.Frame(main_frame)
        options_frame.pack(fill="x", pady=(0, 20))
        
        # Process as Datasheet option
        datasheet_frame = ttk.LabelFrame(options_frame, text="Process as Datasheet", padding="10")
        datasheet_frame.pack(fill="x", pady=(0, 10))
        
        ttk.Label(datasheet_frame, 
                 text="Extract data from datasheet format using coordinates and field mappings.\n"
                      "Use this when your Excel file contains structured data in a datasheet layout.",
                 wraplength=400, justify="left").pack(anchor="w")
        
        # Process as Index option
        index_frame = ttk.LabelFrame(options_frame, text="Process as Index", padding="10")
        index_frame.pack(fill="x")
        
        ttk.Label(index_frame, 
                 text="Generate dictionary from headers and sheet structure.\n"
                      "Use this when your Excel file has headers that define the data structure.",
                 wraplength=400, justify="left").pack(anchor="w")
        
        # Buttons frame
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill="x", pady=(10, 0))
        
        # Result variable
        result = {"choice": None}
        
        def process_as_datasheet():
            result["choice"] = "datasheet"
            dialog.destroy()
        
        def process_as_index():
            result["choice"] = "index"
            dialog.destroy()
        
        def cancel():
            result["choice"] = None
            dialog.destroy()
        
        # Buttons
        ttk.Button(button_frame, text="Process as Datasheet", 
                  command=process_as_datasheet, width=20).pack(side="left", padx=(0, 10))
        ttk.Button(button_frame, text="Process as Index", 
                  command=process_as_index, width=20).pack(side="left", padx=(0, 10))
        ttk.Button(button_frame, text="Cancel", 
                  command=cancel, width=15).pack(side="right")
        
        # Wait for dialog to close
        dialog.wait_window()
        
        return result["choice"]

    def open_change_source_key_dialog(self):
        if not self.selected_coord_for_context:
            return

        coord = self.selected_coord_for_context
        
        # Determine which data type this coordinate belongs to
        coord_type = None
        current_source_key = None
        available_keys = []
        
        for data_type in self.get_all_data_types():
            coordinate_values = self.get_data_type_coordinate_values(data_type)
            if coord in coordinate_values:
                coord_type = data_type
                current_source_key = coordinate_values[coord]
                # Get available keys for this data type
                data = self.get_data_type_data(data_type)
                for key, value in data.items():
                    available_keys = list(value.keys())
                    break
                break
        
        if coord_type is None:
            print(f"Coordinate {coord} not found in any data type coordinate values")
            return

        dialog = Toplevel(self.root)
        dialog.title(f"Change Source Key for {coord}")
        
        # Get cursor position and set dialog position near cursor
        cursor_x = self.root.winfo_pointerx()
        cursor_y = self.root.winfo_pointery()
        
        # Add some offset so dialog doesn't appear exactly at cursor
        dialog_x = cursor_x + 10
        dialog_y = cursor_y + 10
        
        # Ensure dialog doesn't go off screen
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        dialog_width = 400
        dialog_height = 200
        
        if dialog_x + dialog_width > screen_width:
            dialog_x = screen_width - dialog_width - 10
        if dialog_y + dialog_height > screen_height:
            dialog_y = screen_height - dialog_height - 10
            
        dialog.geometry(f"{dialog_width}x{dialog_height}+{dialog_x}+{dialog_y}")
        dialog.transient(self.root)
        dialog.grab_set()

        main_frame = ttk.Frame(dialog, padding="10")
        main_frame.pack(fill=BOTH, expand=True)

        # Current coordinate info
        ttk.Label(main_frame, text=f"Coordinate: {coord}").grid(row=0, column=0, sticky=W, pady=2)
        ttk.Label(main_frame, text=f"Type: {coord_type.upper()}").grid(row=1, column=0, sticky=W, pady=2)
        ttk.Label(main_frame, text=f"Current Source Key: {current_source_key}").grid(row=2, column=0, sticky=W, pady=2)

        # New source key selection
        ttk.Label(main_frame, text="New Source Key:").grid(row=3, column=0, sticky=W, pady=2)
        new_key_combo = ttk.Combobox(main_frame, values=available_keys, state="readonly")
        new_key_combo.grid(row=3, column=1, sticky=EW, padx=5, pady=2)
        new_key_combo.set(current_source_key)  # Set current value as default

        # Buttons
        button_frame = ttk.Frame(dialog)
        button_frame.pack(side=BOTTOM, fill=X, padx=10, pady=10)

        def save_source_key():
            new_key = new_key_combo.get()
            if new_key and new_key != current_source_key:
                # Update coordinate values for the data type
                coord_values = self.get_data_type_coordinate_values(coord_type)
                coord_values[coord] = new_key
                self.set_data_type_coordinate_values(coord_type, coord_values)
                print(f"Changed source key for {coord} from '{current_source_key}' to '{new_key}'")
                
                # Update listboxes
                for i, tab_name in enumerate(self.notebook.tabs()):
                    if self.notebook.tab(i, "text") == 'Coordinates':
                        coordinates_tab = self.notebook.winfo_children()[i]
                        if hasattr(coordinates_tab, 'reinitialize'):
                             coordinates_tab.reinitialize()
                        break
            dialog.destroy()

        ttk.Button(button_frame, text="OK", command=save_source_key).pack(side=RIGHT, padx=5)
        ttk.Button(button_frame, text="Cancel", command=dialog.destroy).pack(side=RIGHT)

        main_frame.columnconfigure(1, weight=1)
        self.root.wait_window(dialog)

    # endregion

    def open_excel_macros_window(self):
        """Opens the Excel Macro Viewer window."""
        macro_window = tk.Toplevel(self.root)
        ExcelMacroViewer(macro_window)

    def open_excel_regex_search_app(self):
        """Opens the Excel Regex Search window."""
        regex_window = tk.Toplevel(self.root)
        ExcelRegexSearchApp(regex_window)
    
    def open_semantic_matcher(self):
        """Opens the Semantic Matcher window."""
        from semantic_matcher import SemanticMatcherApp
        semantic_window = tk.Toplevel(self.root)
        SemanticMatcherApp(semantic_window, self)

    # Semantic similarity methods
    def load_semantic_model_async(self):
        """Load the sentence transformer model in a background thread"""
        def load_model():
            try:
                self.loading_model = True
                print("DEBUG: Importing sentence_transformers...")
                from sentence_transformers import SentenceTransformer
                # Use a lightweight model for faster loading
                self.semantic_model = SentenceTransformer('all-MiniLM-L6-v2')
                self.model_loaded = True
                self.loading_model = False
                
                # Update UI in main thread
                self.root.after(0, self.on_semantic_model_loaded)
            except Exception as e:
                self.loading_model = False
                self.root.after(0, lambda: self.on_semantic_model_error(str(e)))
        
        thread = threading.Thread(target=load_model, daemon=True)
        thread.start()
    
    def on_semantic_model_loaded(self):
        """Called when semantic model is successfully loaded"""
        print("Semantic model loaded successfully!")
        # Update status label if it exists
        if hasattr(self, 'semantic_status_label'):
            self.semantic_status_label.config(text="Semantic model: Ready", foreground="green")
    
    def on_semantic_model_error(self, error_msg):
        """Called when semantic model loading fails"""
        print(f"Error loading semantic model: {error_msg}")
        messagebox.showerror("Error", f"Failed to load semantic model: {error_msg}")
    
    def get_min_score_threshold(self):
        """Ask user for minimum score threshold for automap"""
        from tkinter import simpledialog
        
        # Ask user for minimum score (default 0.3)
        result = simpledialog.askfloat(
            "AutoMap Settings",
            "Enter minimum similarity score (0.0 - 1.0):\n\n" +
            "• 0.1-0.3: Very permissive (more matches, less accurate)\n" +
            "• 0.4-0.6: Balanced (recommended)\n" +
            "• 0.7-0.9: Very strict (fewer matches, more accurate)",
            initialvalue=0.5,
            minvalue=0.0,
            maxvalue=1.0
        )
        
        return result
    
    def compute_semantic_similarity(self, text1: str, text2: str) -> float:
        """Compute cosine similarity between two texts"""
        if not self.model_loaded or not self.semantic_model:
            return 0.0
        
        try:
            # Encode the texts
            embeddings = self.semantic_model.encode([text1, text2])
            
            # Calculate cosine similarity
            similarity = np.dot(embeddings[0], embeddings[1]) / (
                np.linalg.norm(embeddings[0]) * np.linalg.norm(embeddings[1])
            )
            
            return float(similarity)
        except Exception as e:
            print(f"Error computing semantic similarity: {e}")
            return 0.0
    
    def get_cell_values_above_and_left(self, sheet, current_cell):
        """Get the values of the next non-empty cells above and to the left of the current cell"""
        try:
            # Parse current cell coordinate
            import re
            match = re.match(r'([A-Z]+)(\d+)', current_cell)
            if not match:
                return None, None
            
            col_letter, row_num = match.groups()
            col_num = 0
            for char in col_letter:
                col_num = col_num * 26 + (ord(char) - ord('A') + 1)
            row_num = int(row_num)
            
            above_value = None
            left_value = None
            
            # Get value from cells above (skip blank cells)
            if row_num > 1:
                for r in range(row_num - 1, 0, -1):  # Go up from current row
                    above_cell = f"{col_letter}{r}"
                    cell_value = sheet.range(above_cell).value
                    if cell_value and str(cell_value).strip():
                        above_value = str(cell_value).strip()
                        break
            
            # Get value from cells to the left (skip blank cells)
            if col_num > 1:
                for c in range(col_num - 1, 0, -1):  # Go left from current column
                    left_col = ""
                    temp_col = c
                    while temp_col > 0:
                        temp_col -= 1
                        left_col = chr(ord('A') + (temp_col % 26)) + left_col
                        temp_col //= 26
                    
                    left_cell = f"{left_col}{row_num}"
                    cell_value = sheet.range(left_cell).value
                    if cell_value and str(cell_value).strip():
                        left_value = str(cell_value).strip()
                        break
            
            return above_value, left_value
        except Exception as e:
            print(f"Error getting cell values above/left: {e}")
            return None, None
    
    def auto_map_coordinate_semantic(self, data_type, current_coord, min_score=0.3):
        """Version of semantic mapping without user feedback dialogs"""
        print(f"\n=== AUTO_MAP_COORDINATE_SEMANTIC DEBUG ===")
        print(f"Input parameters:")
        print(f"  - data_type: {data_type}")
        print(f"  - current_coord: {current_coord}")
        print(f"  - min_score: {min_score}")
        
        try:
            # Get cell values above and left using the existing method
            sheet = xw.apps.active.books.active.sheets.active
            print(f"  - Active sheet: {sheet.name}")
            
            above_value, left_value = self.get_cell_values_above_and_left(sheet, current_coord)
            print(f"  - Above value: '{above_value}'")
            print(f"  - Left value: '{left_value}'")
            
            # Get available options for this data type
            data = self.get_data_type_data(data_type)
            print(f"  - Data type data retrieved: {data is not None}")
            if data:
                print(f"  - Data keys count: {len(data)}")
                print(f"  - First data key: {list(data.keys())[0] if data else 'None'}")
            
            if not data:
                print("  - ERROR: No data found for data_type")
                return None, 0.0
            
            # Get the options (keys from the first value dict)
            options = []
            for key, value in data.items():
                options = list(value.keys())
                break
            
            print(f"  - Available options: {options}")
            print(f"  - Options count: {len(options)}")
            
            if not options:
                print("  - ERROR: No options found in data")
                return None, 0.0
            
            # Prepare text for semantic comparison
            text_options = []
            if above_value and str(above_value).strip():
                text_options.append(str(above_value).strip())
            if left_value and str(left_value).strip():
                text_options.append(str(left_value).strip())
            
            print(f"  - Text options for comparison: {text_options}")
            print(f"  - Text options count: {len(text_options)}")
            
            if not text_options:
                print("  - ERROR: No text options available for comparison")
                return None, 0.0
            
            # Check if semantic model is available
            if not hasattr(self, 'semantic_model') or self.semantic_model is None:
                print("  - ERROR: Semantic model not available")
                return None, 0.0
            
            print(f"  - Semantic model available: {self.semantic_model is not None}")
            
            # Encode options and text for comparison
            print("  - Encoding options...")
            option_embeddings = self.semantic_model.encode(options)
            print(f"  - Option embeddings shape: {option_embeddings.shape}")
            
            best_match = None
            best_score = 0
            best_header = None
            
            print("  - Computing similarities...")
            for i, text in enumerate(text_options):
                print(f"    Processing text {i+1}: '{text}'")
                text_embedding = self.semantic_model.encode([text])
                similarities = np.dot(text_embedding, option_embeddings.T).flatten()
                
                print(f"    Similarities: {similarities}")
                max_idx = np.argmax(similarities)
                score = similarities[max_idx]
                best_option = options[max_idx]
                
                print(f"    Best match: '{best_option}' with score: {score:.4f}")
                
                if score > best_score:
                    best_score = score
                    best_match = best_option
                    best_header = text
                    print(f"    -> New best match!")
            
            print(f"  - Final best match: '{best_match}'")
            print(f"  - Final best score: {best_score:.4f}")
            print(f"  - Final best header: '{best_header}'")
            print(f"  - Score threshold: {min_score}")
            print(f"  - Meets threshold: {best_score > min_score}")
            
            # Return best match, score, and header
            if best_match and best_score > min_score:
                print(f"  - RESULT: SUCCESS - Returning match '{best_match}' with score {best_score:.4f}")
                return best_match, best_score, best_header
            else:
                print(f"  - RESULT: FAILED - Score {best_score:.4f} below threshold {min_score}")
                return None, best_score, best_header
                
        except Exception as e:
            print(f"  - ERROR in semantic mapping: {e}")
            import traceback
            print(f"  - Traceback: {traceback.format_exc()}")
            return None, 0.0

    def automap_range(self, data_type, full_selection, combo_box, coord_entry, listbox, min_score=0.3):
        """Iterate through a range of cells and perform semantic mapping"""
        try:
            # Parse the range
            clean_selection = full_selection.replace('$', '')
            
            # Extract all cell addresses from the range
            cell_addresses = self.extract_cell_addresses_from_range(clean_selection)
            print('debug: cell_addresses', cell_addresses)
            # Process all cells using the unified processor
            successful_mappings, mapping_results = self.process_cells_for_automap(data_type, cell_addresses, min_score)
            
            # Update the listbox to show all new mappings
            if hasattr(self, f"{data_type}_listbox"):
                listbox_obj = getattr(self, f"{data_type}_listbox")
                self.update_single_listbox(data_type, listbox_obj)
            
            # Show results summary
            summary = f"AutoMap Range Results:\n\n"
            summary += f"Range: {clean_selection}\n"
            summary += f"Successful mappings: {successful_mappings}\n"
            summary += f"Total cells processed: {len(mapping_results)}\n"
            summary += f"Minimum score threshold: {min_score}\n\n"
            
            if mapping_results:
                summary += "Detailed Results:\n" + "\n".join(mapping_results[:10])  # Show first 10
                if len(mapping_results) > 10:
                    summary += f"\n... and {len(mapping_results) - 10} more results"
            
            messagebox.showinfo("AutoMap Results", summary)
            print(f"AutoMap completed: {successful_mappings} successful mappings")
            
        except Exception as e:
            print(f"Error in automap range: {e}")

    def automap_noncontiguous(self, data_type, full_selection, combo_box, coord_entry, listbox, min_score=0.3):
        """Process non-contiguous selections (multiple ranges separated by commas)"""
        try:
            # Parse the non-contiguous selection
            clean_selection = full_selection.replace('$', '')
            ranges = clean_selection.split(',')
            
            total_successful_mappings = 0
            all_mapping_results = []
            
            print(f"Processing {len(ranges)} non-contiguous ranges...")
            
            for i, range_addr in enumerate(ranges):
                range_addr = range_addr.strip()
                print(f"Processing range {i+1}/{len(ranges)}: {range_addr}")
                
                if ':' in range_addr:
                    # This is a range (e.g., A1:B5)
                    cell_addresses = self.extract_cell_addresses_from_range(range_addr)
                    successful_mappings, mapping_results = self.process_cells_for_automap(data_type, cell_addresses, min_score)
                else:
                    # This is a single cell (e.g., C3)
                    successful_mappings, mapping_results = self.process_cells_for_automap(data_type, [range_addr], min_score)
                
                total_successful_mappings += successful_mappings
                all_mapping_results.extend(mapping_results)
                print(f"Range {i+1} completed: {successful_mappings} mappings")
            
            # Update the listbox to show all new mappings
            if hasattr(self, f"{data_type}_listbox"):
                listbox_obj = getattr(self, f"{data_type}_listbox")
                self.update_single_listbox(data_type, listbox_obj)
            
            # Show results summary
            summary = f"AutoMap Non-Contiguous Results:\n\n"
            summary += f"Selection: {clean_selection}\n"
            summary += f"Ranges processed: {len(ranges)}\n"
            summary += f"Successful mappings: {total_successful_mappings}\n"
            summary += f"Total cells processed: {len(all_mapping_results)}\n"
            summary += f"Minimum score threshold: {min_score}\n\n"
            
            if all_mapping_results:
                summary += "Detailed Results:\n" + "\n".join(all_mapping_results[:15])  # Show first 15
                if len(all_mapping_results) > 15:
                    summary += f"\n... and {len(all_mapping_results) - 15} more results"
            
            messagebox.showinfo("AutoMap Results", summary)
            print(f"Non-contiguous AutoMap completed: {total_successful_mappings} total successful mappings")
            
        except Exception as e:
            print(f"Error in automap non-contiguous: {e}")



    def update_single_listbox(self, data_type, listbox):
        """Update a single listbox for a specific data type"""
        listbox.delete(0, tk.END)
        coordinate_values = self.get_data_type_coordinate_values(data_type)
        for key, value in coordinate_values.items():
            coord_display = f"{key}: {value}"
            if key in self.coordinate_conversions:
                conv = self.coordinate_conversions[key]
                coord_display += f" [{conv.get('in_unit', '?')}->{conv.get('out_unit', '?')}]"
            if key in self.coordinate_combinations:
                combo = self.coordinate_combinations[key]
                coord_display += f" [Combines: {', '.join(combo.get('combines', []))} ({combo.get('operation', 'add')})]"
            listbox.insert(tk.END, coord_display)

    def process_cells_for_automap(self, data_type, cell_addresses, min_score=0.3):
        """Unified function to process any collection of cells for automapping"""
        print(f"\n=== PROCESS_CELLS_FOR_AUTOMAP DEBUG ===")
        print(f"Input parameters:")
        print(f"  - data_type: {data_type}")
        print(f"  - cell_addresses: {cell_addresses}")
        print(f"  - min_score: {min_score}")
        print(f"  - Total cells to process: {len(cell_addresses)}")
        
        try:
            # Get the sheet
            sheet = xw.apps.active.books.active.sheets.active
            print(f"  - Active sheet: {sheet.name}")
            
            # Handle merged cells and iterate through unique cells
            processed_cells = set()
            successful_mappings = 0
            mapping_results = []
            
            # Print initial cell count
            print(f"Starting automap with {len(cell_addresses)} input cells...")
            
            for i, cell_addr in enumerate(cell_addresses):
                print(f"\n--- Processing cell {i+1}/{len(cell_addresses)}: {cell_addr} ---")
                cell_addr = cell_addr.strip()
                if not cell_addr:
                    print(f"  - Skipping empty cell address")
                    continue
                    
                # Skip if we've already processed this cell (due to merged cells)
                if cell_addr in processed_cells:
                    print(f"  - Skipping already processed cell: {cell_addr}")
                    continue
                
                # Check if this cell is part of a merged range
                merged_range = None
                try:
                    print(f"  - Checking for merged cells...")
                    # Try a different approach to check for merged cells
                    cell_range = sheet.range(cell_addr)
                    if hasattr(cell_range.api, 'MergeCells') and cell_range.api.MergeCells:
                        # This cell is part of a merged range
                        merged_range = cell_range.api.MergeArea.Address.replace('$', '')
                        print(f"  - Found merged range: {merged_range}")
                        
                        # Get the top-left cell of the merged range
                        top_left_cell = merged_range.split(':')[0]
                        print(f"  - Top-left cell of merged range: {top_left_cell}")
                        
                        # If this cell is not the top-left cell, skip it
                        if cell_addr != top_left_cell:
                            print(f"  - Skipping {cell_addr} (not top-left of merged range {merged_range})")
                            processed_cells.add(cell_addr)
                            continue
                        
                        # Add all cells in the merged range to processed set to avoid reprocessing
                        merge_start, merge_end = merged_range.split(':')
                        merge_range_obj = sheet.range(f"{merge_start}:{merge_end}")
                        # Use shape to iterate through merged range
                        merge_rows = merge_range_obj.shape[0]
                        merge_cols = merge_range_obj.shape[1]
                        print(f"  - Merged range dimensions: {merge_rows}x{merge_cols}")
                        for mr in range(merge_rows):
                            for mc in range(merge_cols):
                                merge_cell = merge_range_obj.offset(mr, mc).resize(1, 1)
                                processed_cells.add(merge_cell.address.replace('$', ''))
                        print(f"  - Added {merge_rows * merge_cols} cells from merged range to processed set")
                    else:
                        print(f"  - No merged cells found for {cell_addr}")
                except Exception as e:
                    # If merged cells check fails, continue normally
                    print(f"  - Merged cells check failed: {e}")
                    pass
            
                # Use the top-left cell of merged range or the single cell
                target_cell = merged_range.split(':')[0] if merged_range else cell_addr
                print(f"  - Target cell for processing: {target_cell}")
                
                # For auto mapping, we don't need to check if the selected cell has a value
                # because we're mapping based on headers (above/left), not the data values
                print(f"  - Processing cell for auto mapping (value check skipped)")
                
                # Perform semantic mapping for this cell
                print(f"  - Calling auto_map_coordinate_semantic...")
                best_match, score, header = self.auto_map_coordinate_semantic(data_type, target_cell, min_score)
                print(f"  - Semantic mapping result:")
                print(f"    - best_match: {best_match}")
                print(f"    - score: {score}")
                print(f"    - header: {header}")
                
                # Record the result
                if best_match:
                    print(f"  - SUCCESS: Adding mapping for {target_cell}")
                    # Add the mapping
                    coordinate_values = self.get_data_type_coordinate_values(data_type)
                    print(f"  - Current coordinate values count: {len(coordinate_values)}")
                    coordinate_values[target_cell] = best_match
                    self.set_data_type_coordinate_values(data_type, coordinate_values)
                    successful_mappings += 1
                    header_display = f"'{header}'" if header else "None"
                    mapping_results.append(f"✓ {target_cell} → {best_match} (score: {score:.3f}, header: {header_display})")
                    print(f"  - Mapping added successfully")
                else:
                    print(f"  - FAILED: No match found for {target_cell}")
                    header_display = f"'{header}'" if header else "None"
                    mapping_results.append(f"✗ {target_cell} → No match (score: {score:.3f}, header: {header_display})")
                
                processed_cells.add(cell_addr)
                print(f"  - Added {cell_addr} to processed cells set")
            
            # Print total unique cells processed (merged cells count as one)
            print(f"\n=== AUTOMAP SUMMARY ===")
            print(f"Total unique cells processed: {len(processed_cells)}")
            print(f"Successful mappings: {successful_mappings}")
            print(f"Failed mappings: {len(mapping_results) - successful_mappings}")
            print(f"Success rate: {(successful_mappings/len(processed_cells)*100):.1f}%" if processed_cells else "N/A")
            
            return successful_mappings, mapping_results
            
        except Exception as e:
            print(f"  - ERROR in process_cells_for_automap: {e}")
            import traceback
            print(f"  - Traceback: {traceback.format_exc()}")
            return 0, []

    def extract_cell_addresses_from_range(self, range_addr):
        """Extract all individual cell addresses from a range address"""
        try:
            # Get the sheet
            sheet = xw.apps.active.books.active.sheets.active
            
            # Get the range object
            range_obj = sheet.range(range_addr)
            
            # Get the dimensions of the range
            rows = range_obj.shape[0]
            cols = range_obj.shape[1]
            
            cell_addresses = []
            
            # Iterate through each cell in the range
            for row in range(rows):
                for col in range(cols):
                    # Get the current cell
                    cell = range_obj.offset(row, col).resize(1, 1)
                    cell_address = cell.address.replace('$', '')
                    cell_addresses.append(cell_address)
            
            return cell_addresses
            
        except Exception as e:
            print(f"Error extracting cell addresses from range {range_addr}: {e}")
            return []


if __name__ == "__main__":
    print("DEBUG: Starting main execution...")
    print("DEBUG: Creating tkinter root window...")
    root = tk.Tk()
    print("DEBUG: Creating DatasheetGeneratorApp instance...")
    app = DatasheetGeneratorApp(root)
    print("DEBUG: Starting main event loop...")
    root.mainloop()
    print("DEBUG: Main event loop ended.")

