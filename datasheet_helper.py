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
        
        # Settings file tracking
        self.current_settings_file = None  # Track the currently loaded/saved settings file
        self.last_settings_file_path = "last_settings_file.txt"  # File to store the last settings file path
        
        # History tracking for dialog inputs (keep last 7 entries)
        self.split_keys_history = []  # Recent delimiters used in Split Keys
        self.modify_keys_history = []  # Recent transformation codes used in Modify Keys
        
        # Centralized data source definitions
        print("DEBUG: Setting up data source definitions...")
        self.data_sources = {}
        
        # Widget references for each data source
        self.data_source_widgets = {}  # Structure: {data_source: {'frame': widget, 'config_notebook': widget, 'combo': widget, 'listbox': widget}}
        
        # Centralized destination definitions
        print("DEBUG: Setting up destination definitions...")
        self.destinations = {}
        
        # Widget references for each destination
        self.destination_widgets = {}  # Structure: {destination: {'frame': widget, 'entries': {...}}}



        # Semantic similarity model
        print("DEBUG: Initializing semantic model attributes...")
        self.semantic_model = None
        self.model_loaded = False
        self.loading_model = False
        print(f"DEBUG: Initial state - semantic_model: {self.semantic_model}, model_loaded: {self.model_loaded}, loading_model: {self.loading_model}")
        
        # Safety check: If model_loaded is True but semantic_model is None (from bad state restore), reset the flags
        if self.model_loaded and self.semantic_model is None:
            print("DEBUG: WARNING - Found inconsistent state (model_loaded=True but semantic_model=None). Resetting flags.")
            self.model_loaded = False
            self.loading_model = False
        
        # Centralized Excel selection monitoring
        self.current_excel_selection = None
        self.coordinate_update_callbacks = []  # List of callback functions for each tab
        self.excel_monitoring_active = False
        self.monitor_excel_selection_id = None  # Store after() callback ID to cancel it
        
        # Initialize datasheets attribute
        self.destination_datasheet = None
        
        # Create a default destination if none exist (for backward compatibility)
        if not self.destinations:
            default_dest = "Default"
            self.destinations[default_dest] = {
                'path': '',
                'datasheet_coord': '',
                'ds_str': '',
                'rows_per_sheet': 1,
                'sig_figs': 4,
                'rounding_tolerance': 1e-2
            }
            print(f"DEBUG: Created default destination: {default_dest}")

        print("DEBUG: About to create widgets...")
        self.create_widgets()
        print("DEBUG: DatasheetGeneratorApp initialization completed!")
        
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        # Start centralized Excel monitoring
        self.start_excel_monitoring()
        
        # Auto-load the last settings file if it exists
        self.auto_load_last_settings()
    

    
    def update_instance_attributes_from_entries(self):
        """Update instance attributes from entry widget values"""
        # Clean up destroyed widgets from entries list
        self.cleanup_entries()
        
        for entry, variable in self.entries:
            value = entry.get()
            
            # Handle different data sources
            if variable == 'rows_per_sheet':
                self.rows_per_sheet = int(value)
            elif variable == 'sig_figs':
                self.sig_figs = int(value)
            elif variable == 'rounding_tolerance':
                self.rounding_tolerance = float(value)
            elif variable == 'datasheet_coord':
                self.datasheet_coord = value
            elif variable == 'ds_str':
                self.ds_str = value
            elif variable == 'datasheets':
                self.destination_datasheet = value
    
    def cleanup_entries(self):
        """Remove destroyed widgets from the entries list"""
        if not hasattr(self, 'entries'):
            return
            
        # Create a new list with only existing widgets
        valid_entries = []
        for entry, variable in self.entries:
            try:
                # Check if the widget still exists
                if entry.winfo_exists():
                    valid_entries.append((entry, variable))
            except tk.TclError:
                # Widget has been destroyed, skip it
                continue
        
        # Update the entries list
        self.entries = valid_entries
    
    def update_combo_boxes(self):
        """Update combo box values in the data source tabs"""
        if not hasattr(self, 'data_sources_notebook'):
            return
        
        # Get all data sources
        data_sources = self.get_all_data_sources()
        
        for data_source in data_sources:
            try:
                # Find the tab for this data source
                tab_found = False
                for i in range(self.data_sources_notebook.index("end")):
                    tab_text = self.data_sources_notebook.tab(i, "text")
                    base_tab_text = self.strip_tab_indicators(tab_text)
                    
                    if base_tab_text == data_source:
                        tab_found = True
                        # Get the tab content
                        tab_content = self.data_sources_notebook.nametowidget(self.data_sources_notebook.tabs()[i])
                        
                        # Find and update combo boxes in this tab
                        self.update_combo_boxes_in_widget(tab_content, data_source)
                        break
                
                if not tab_found:
                    print(f"DEBUG: Tab not found for data source: {data_source}")
                    
            except Exception as e:
                print(f"Error updating combo boxes for {data_source}: {e}")
    
    def update_combo_boxes_in_widget(self, widget, data_source):
        """Recursively find and update combo boxes in a widget and its children"""
        try:
            # Check if this widget is a combobox
            if isinstance(widget, ttk.Combobox):
                # Update the combobox values based on the data source
                if hasattr(widget, 'data_source') and widget.data_source == data_source:
                    # This is a source sheet combobox - update with available sheets
                    file_path = self.data_sources[data_source]['path']
                    if file_path and os.path.exists(file_path):
                        try:
                            # Get available sheet names from the Excel file using the same method as Configure
                            sheet_names = self.get_sheet_names_from_file(file_path)
                            
                            # Update combobox values
                            widget['values'] = sheet_names
                            
                            # Set current value if it's still valid
                            current_value = widget.get()
                            if current_value not in sheet_names and sheet_names:
                                widget.set(sheet_names[0])  # Set to first sheet if current is invalid
                                
                        except Exception as e:
                            print(f"Error updating sheet names for {data_source}: {e}")
                            widget['values'] = []
            
            # Recursively check children
            for child in widget.winfo_children():
                self.update_combo_boxes_in_widget(child, data_source)
                
        except Exception as e:
            print(f"Error in update_combo_boxes_in_widget: {e}")
    
    def get_sheet_names_from_file(self, file_path):
        """Get sheet names from any Excel file using the same method as Configure button"""
        if not file_path or not os.path.exists(file_path):
            return []
        
        try:
            # Use openpyxl for better performance and reliability (same as configure_source_data_dialog)
            wb = openpyxl.load_workbook(file_path, read_only=True)
            sheet_names = wb.sheetnames
            wb.close()
            return sheet_names
        except Exception as e:
            print(f"Error getting sheet names from {file_path}: {e}")
            return []
    
    def update_coordinates_combo_box(self, data_source):
        """Update the coordinates section combo box with data from data_source_data"""
        try:
            # Get the combo box for this data source
            combo = self.data_source_widgets.get(data_source, {}).get('combo')
            if combo:
                
                # Check if the widget still exists
                if not combo.winfo_exists():
                    print(f"Warning: Combo box for {data_source} no longer exists")
                    return
                
                # Get the data for this data source
                data = self.data_sources[data_source]['data']
                values = []
                
                if data:
                    # Get keys from the first entry (same logic as get_combo_values in create_coordinates_tab)
                    first_entry = list(data.values())[0]
                    if isinstance(first_entry, dict):
                        values = list(first_entry.keys())
                
                # Update the combo box values
                combo['values'] = values
                
                # If current value is not in the new values, clear it
                current_value = combo.get()
                if current_value and current_value not in values:
                    combo.set('')
                
                print(f"Updated coordinates combo box for {data_source} with {len(values)} values")
            else:
                print(f"Warning: Combo box not found for data source {data_source}")
                
        except Exception as e:
            print(f"Error updating coordinates combo box for {data_source}: {e}")
    
    def update_all_coordinates_combo_boxes(self):
        """Update all coordinates section combo boxes with current data"""
        for data_source in self.get_all_data_sources():
            self.update_coordinates_combo_box(data_source)
        print("Updated all coordinates combo boxes")
    
    
    # region Centralized Excel Selection Monitoring
    
    def start_excel_monitoring(self):
        """Start the centralized Excel selection monitoring"""
        if not self.excel_monitoring_active:
            self.excel_monitoring_active = True
            self.monitor_excel_selection()
    
    def stop_excel_monitoring(self):
        """Stop the centralized Excel selection monitoring"""
        self.excel_monitoring_active = False
        # Cancel any pending callback to prevent it from executing
        if self.monitor_excel_selection_id and self.root:
            try:
                self.root.after_cancel(self.monitor_excel_selection_id)
                self.monitor_excel_selection_id = None
            except Exception as e:
                print(f"Error canceling Excel monitoring callback: {e}")
    
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
            if self.destination_datasheet:
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
            self.monitor_excel_selection_id = self.root.after(200, self.monitor_excel_selection)
    
    # endregion
    
    
    def get_all_data_sources(self):
        """Get all available data sources"""
        return list(self.data_sources.keys())
    
    def get_all_destinations(self):
        """Get all destination names"""
        return list(self.destinations.keys())
    
    def get_current_destination(self):
        """Get the current destination (currently selected tab)"""
        if hasattr(self, 'destinations_notebook'):
            try:
                selected_tab = self.destinations_notebook.select()
                if selected_tab:
                    tab_text = self.destinations_notebook.tab(selected_tab, "text")
                    if tab_text in self.destinations:
                        return tab_text
            except Exception as e:
                print(f"Error getting selected destination tab: {e}")
        
        # Fallback to first destination if no tab is selected
        destinations = self.get_all_destinations()
        if destinations:
            return destinations[0]
        return None
    
    def get_primary_data_source(self):
        """Get the primary data source (currently selected tab)"""
        if hasattr(self, 'data_sources_notebook'):
            try:
                # Get the currently selected tab
                selected_tab = self.data_sources_notebook.select()
                if selected_tab:
                    # Get the tab text which contains the data source name
                    tab_text = self.data_sources_notebook.tab(selected_tab, "text")
                    # Remove checkmark if present
                    data_source = self.strip_tab_indicators(tab_text).strip()
                    # Verify this data source exists
                    if data_source in self.data_sources:
                        return data_source
            except Exception as e:
                print(f"Error getting selected tab: {e}")
        
        # Fallback to first data source if no tab is selected or error occurs
        return list(self.data_sources.keys())[0] if self.data_sources else None
    
    
    def on_tab_changed(self, event):
        """Handle tab change event to update primary data source"""
        primary_data_source = self.get_primary_data_source()
        if primary_data_source:
            config = self.data_sources[primary_data_source]
            name = primary_data_source
            print(f"Primary data source changed to: {name} ({primary_data_source})")
    
    def select_data_source_tab(self, data_source_name):
        """Select the tab for the given data source name if it exists"""
        if not hasattr(self, 'data_sources_notebook'):
            return
        try:
            for tab_id in self.data_sources_notebook.tabs():
                tab_text = self.data_sources_notebook.tab(tab_id, "text")
                normalized_text = self.strip_tab_indicators(tab_text).strip()
                if normalized_text == data_source_name:
                    self.data_sources_notebook.select(tab_id)
                    break
        except Exception as e:
            print(f"Error selecting data source tab '{data_source_name}': {e}")
    
    def open_data_source_switcher(self):
        """Open a dialog listing all datasources; selecting one switches to that tab"""
        dialog = Toplevel(self.root)
        dialog.title("Datasources")
        dialog.geometry("300x400")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Get all data sources and store them
        all_data_sources = self.get_all_data_sources()
        
        # Search entry box
        search_frame = Frame(dialog)
        search_frame.pack(fill=X, padx=10, pady=(10, 5))
        Label(search_frame, text="Search:").pack(side=LEFT, padx=(0, 5))
        search_entry = Entry(search_frame)
        search_entry.pack(side=LEFT, fill=X, expand=True)
        search_entry.focus_set()
        
        # Listbox with all data sources
        listbox = tk.Listbox(dialog)
        for name in all_data_sources:
            listbox.insert(END, name)
        listbox.pack(fill=BOTH, expand=True, padx=10, pady=(5, 10))
        
        def filter_listbox(event=None):
            """Filter the listbox based on search entry"""
            search_term = search_entry.get().lower()
            listbox.delete(0, END)
            for name in all_data_sources:
                if search_term in name.lower():
                    listbox.insert(END, name)
            # Select first item if available
            if listbox.size() > 0:
                listbox.selection_set(0)
                listbox.activate(0)
        
        # Bind search entry to filter function
        search_entry.bind('<KeyRelease>', filter_listbox)
        
        def on_select(event=None):
            selection = listbox.curselection()
            if selection:
                name = listbox.get(selection[0])
                self.select_data_source_tab(name)
                dialog.destroy()
        
        # Double-click selects
        listbox.bind("<Double-Button-1>", on_select)
        listbox.bind('<Return>', on_select)
        
        # Buttons
        btn_frame = Frame(dialog)
        btn_frame.pack(fill=X, padx=10, pady=(0, 10))
        ttk.Button(btn_frame, text="Open", command=on_select, width=10).pack(side=RIGHT, padx=5)
        ttk.Button(btn_frame, text="Cancel", command=dialog.destroy, width=10).pack(side=RIGHT, padx=5)
        
        # Center over parent
        center_window_over_parent(dialog)
    
    def add_data_source_from_entry(self, event=None):
        """Add a new data source from the text entry box"""
        data_source_id = self.add_data_source_entry.get().strip()
        
        # Validation
        if not data_source_id:
            messagebox.showerror("Error", "data source ID is required")
            return
        
        # Check if data source already exists
        if data_source_id in self.data_sources:
            messagebox.showerror("Error", f"data source '{data_source_id}' already exists")
            return
        
        # Create the new data source configuration
        new_config = {
            'headers': [],
            'coordinate_values': {},  # Will be structured as {destination: {coord: value}}
            'selected_sheets': None,
            'path': '',  # File path stored directly in config
            'top_tag': '',  # Default top tag
            'source_sheet_name': 'TEMPLATE',  # Default source sheet name
            'partial_match': False,  # Default to exact matching
            'data': {}  # Data will be stored as nested dictionaries here
        }
        
        # Add to data sources
        self.add_data_source(data_source_id, new_config)
        
        # Clear the entry box
        self.add_data_source_entry.delete(0, tk.END)
        
        # Add just the new tab without refreshing everything
        self.add_single_data_source_tab(data_source_id, new_config)
        
        messagebox.showinfo("Success", f"data source '{data_source_id}' created successfully!")

    def add_destination_from_entry(self, event=None):
        """Add a new destination from the text entry box"""
        destination_id = self.add_destination_entry.get().strip()
        
        # Validation
        if not destination_id:
            messagebox.showerror("Error", "Destination ID is required")
            return
        
        # Check if destination already exists
        if destination_id in self.destinations:
            messagebox.showerror("Error", f"Destination '{destination_id}' already exists")
            return
        
        # Create the new destination configuration
        new_config = {
            'path': '',  # File path for destination datasheet
            'datasheet_coord': '',
            'ds_str': '',  # Prefix
            'rows_per_sheet': 1,
            'sig_figs': 4,
            'rounding_tolerance': 1e-2
        }
        
        # Add to destinations
        self.add_destination(destination_id, new_config)
        
        # Clear the entry box
        self.add_destination_entry.delete(0, tk.END)
        
        # Note: add_destination() already calls refresh_destinations_notebook() which adds the tab
        # No need to call add_single_destination_tab() separately
        
        messagebox.showinfo("Success", f"Destination '{destination_id}' created successfully!")

    def add_single_data_source_tab(self, data_source, config):
        """Add a single new data source tab without affecting existing tabs or entries"""
        if not hasattr(self, 'data_sources_notebook'):
            return
        
        # Create main tab for this data source
        data_source_tab = ttk.Frame(self.data_sources_notebook)
        self.data_sources_notebook.add(data_source_tab, text=data_source)
        
        # Create the data source tab content
        self.create_data_source_tab_content(data_source_tab, data_source, config)
        
        # Update tab indicators (this doesn't affect entries)
        self.update_all_tab_indicators()
        
        print(f"Added new tab for data source: {data_source}")


    def add_data_source(self, data_source, config):
        """Add a new data source configuration"""
        # Ensure the config has all required fields with defaults
        default_config = {
            'headers': [],
            'coordinate_values': {},
            'selected_sheets': None,
            'top_tag': '',
            'source_sheet_name': 'TEMPLATE',
            'partial_match': False,
            'data': {},
            'tag_filters': [],
            'transform_data_source': None,
            'transform_key': None,
            'transformation_code': '',
            'coordinate_conversions': {},
            'coordinate_combinations': {},
            'extraction_settings': {
                'init_tag_coord': '',
                'init_coords_to_fields': {},
                'tags_per_sheet': 1,
                'selected_sheets': []
            }
        }
        default_config.update(config)
        
        # Add to the centralized data sources dictionary
        self.data_sources[data_source] = default_config
        
        # Initialize path in config if not present
        if 'path' not in self.data_sources[data_source]:
            self.data_sources[data_source]['path'] = ''
        
        # Initialize coordinate maps for all destinations with this new data source
        self.initialize_coordinate_maps_for_all_combinations()
        
        # Refresh the GUI to include the new data source
        if hasattr(self, 'root') and self.root:
            self.refresh_data_source_frames()
    
    def add_destination(self, destination, config):
        """Add a new destination configuration"""
        # Ensure the config has all required fields with defaults
        default_config = {
            'path': '',
            'datasheet_coord': '',
            'ds_str': '',
            'rows_per_sheet': 1,
            'sig_figs': 4,
            'rounding_tolerance': 1e-2
        }
        default_config.update(config)
        
        # Add to the centralized destinations dictionary
        self.destinations[destination] = default_config
        
        # Initialize coordinate maps for all data sources with this new destination
        self.initialize_coordinate_maps_for_all_combinations()
        
        # Refresh the GUI to include the new destination
        if hasattr(self, 'root') and self.root:
            if hasattr(self, 'destinations_notebook'):
                self.refresh_destinations_notebook()
    
    def initialize_coordinate_maps_for_all_combinations(self):
        """Initialize coordinate maps for all data source * destination combinations.
        Ensures that every data source has a coordinate_values entry for every destination.
        Also handles migration from old format (flat dict) to new format (nested by destination)."""
        all_data_sources = self.get_all_data_sources()
        all_destinations = self.get_all_destinations()
        
        # If no destinations exist, create a default one (directly to avoid recursion)
        if not all_destinations:
            default_dest = "Default"
            if default_dest not in self.destinations:
                default_config = {
                    'path': '',
                    'datasheet_coord': '',
                    'ds_str': '',
                    'rows_per_sheet': 1,
                    'sig_figs': 4,
                    'rounding_tolerance': 1e-2
                }
                self.destinations[default_dest] = default_config
                # Refresh GUI if available
                if hasattr(self, 'root') and self.root:
                    if hasattr(self, 'destinations_notebook'):
                        self.refresh_destinations_notebook()
            all_destinations = self.get_all_destinations()
        
        for data_source in all_data_sources:
            # Ensure coordinate_values exists and is a dict
            if 'coordinate_values' not in self.data_sources[data_source]:
                self.data_sources[data_source]['coordinate_values'] = {}
            elif not isinstance(self.data_sources[data_source]['coordinate_values'], dict):
                # Handle corrupted data - reset to empty dict
                self.data_sources[data_source]['coordinate_values'] = {}
            
            coord_vals = self.data_sources[data_source]['coordinate_values']
            
            # Check if coordinate_values is in old format (flat dict with coordinate keys)
            # New format has coordinate_values[destination] = {coord: value}
            if coord_vals:
                # Check if any key matches a destination name (new format)
                is_nested = any(dest in coord_vals for dest in all_destinations)
                
                if not is_nested and coord_vals:
                    # Likely old format - check if keys look like coordinates (A1, B2, etc.)
                    sample_key = list(coord_vals.keys())[0] if coord_vals else None
                    if sample_key and (len(sample_key) <= 4 and sample_key[0].isalpha()):
                        # Old format detected - migrate to new format using first/default destination
                        default_dest = all_destinations[0]
                        old_coords = coord_vals.copy()
                        self.data_sources[data_source]['coordinate_values'] = {
                            default_dest: old_coords
                        }
                        print(f"Migrated coordinate_values for '{data_source}' from old format to '{default_dest}' destination")
                        coord_vals = self.data_sources[data_source]['coordinate_values']
            
            # Initialize coordinate_values for each destination (create empty dict if missing)
            for destination in all_destinations:
                if destination not in coord_vals:
                    coord_vals[destination] = {}

    def add_single_destination_tab(self, destination, config):
        """Add a single new destination tab without affecting existing tabs or entries"""
        if not hasattr(self, 'destinations_notebook'):
            return
        
        # Create main tab for this destination
        destination_tab = ttk.Frame(self.destinations_notebook)
        self.destinations_notebook.add(destination_tab, text=destination)
        
        # Create the destination tab content
        self.create_destination_tab_content(destination_tab, destination, config)
        
        print(f"Added new tab for destination: {destination}")

    
    def create_data_source_frames(self, parent):
        """Dynamically create UI frames for all data sources"""
        print("DEBUG: Starting create_data_source_frames...")
        # Add a label for the data sources section
        sources_label = tk.Label(parent, text="DATA SOURCES", font=("Arial", 10, "bold"), fg="green")
        sources_label.pack(anchor=tk.W, padx=5, pady=(5,0))
        
        print("DEBUG: Creating frames for each data source...")
        for data_source in self.get_all_data_sources():
            config = self.data_sources[data_source]
            name = data_source
            
            # Create frame for this data source
            frame = tk.Frame(parent)
            frame.pack(fill=tk.X, pady=5)
            
            # Label
            label = tk.Label(frame, text=f"{name} (Source)", width=30)
            label.pack(side=tk.LEFT, padx=5)
            
            # Entry
            entry = tk.Entry(frame)
            entry._data_source = data_source  # Store data source name for repopulation
            entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
            
            # Buttons frame
            buttons = tk.Frame(frame)
            buttons.pack(side=tk.RIGHT)
            
            # Browse button
            tk.Button(buttons, text="Browse",
                      command=lambda e=entry, ds=data_source: self.browse_data_source(e, ds)).pack(side=tk.LEFT, padx=2)
            
            # Open button
            tk.Button(buttons, text="Open",
                      command=lambda e=entry: self.open_file(e)).pack(side=tk.LEFT, padx=2)
            
            # Configure button
            tk.Button(buttons, text="Configure",
                      command=lambda n=name: self.configure(n)).pack(side=tk.LEFT, padx=2)
            
            # Generate button
            tk.Button(buttons, text="Generate",
                      command=lambda dt=data_source: self.generate_data_source(dt)).pack(side=tk.LEFT, padx=2)
            
            # View button
            tk.Button(buttons, text="View",
                      command=lambda ds=data_source: self.view_data(data_source=ds)).pack(side=tk.LEFT, padx=2)
            
            # Store reference to frame for potential updates
            if data_source not in self.data_source_widgets:
                self.data_source_widgets[data_source] = {}
            self.data_source_widgets[data_source]['frame'] = frame
        print("DEBUG: create_data_source_frames completed!")
    
    def create_data_sources_notebook(self, parent):
        """Create a notebook with tabs for each data source, each containing nested configuration tabs"""
        print("DEBUG: Starting create_data_sources_notebook...")
        
        
        # Add a label and button for the data sources section
        sources_frame = ttk.Frame(parent)
        sources_frame.pack(fill="x", padx=5, pady=(5,0))
        
        sources_label = tk.Label(sources_frame, text="DATA SOURCES", font=("Arial", 10, "bold"), fg="green")
        sources_label.pack(side="left")
        
        # Add search bar for data sources
        search_frame = ttk.Frame(sources_frame)
        search_frame.pack(side="left", padx=(10, 0))
        
        self.data_source_search_entry = tk.Entry(search_frame, width=25, foreground='gray')
        self.data_source_search_entry.pack(side="left", padx=(0, 5))
        self.data_source_search_entry.bind('<KeyRelease>', self.search_data_sources)
        self.data_source_search_entry.insert(0, "Search keys/values...")
        self.data_source_search_entry.bind('<FocusIn>', lambda e: self.clear_search_placeholder())
        self.data_source_search_entry.bind('<FocusOut>', lambda e: self.restore_search_placeholder())
        
        # Clear search button
        clear_search_btn = ttk.Button(search_frame, text="✕", width=3, 
                                     command=self.clear_data_source_search)
        clear_search_btn.pack(side="left")
        
        # Button to open a list of datasources and switch tabs
        switch_btn = ttk.Button(search_frame, text="Switch…", width=9, 
                                command=self.open_data_source_switcher)
        switch_btn.pack(side="left", padx=(5, 0))
        
        # Add text entry and button to create new data source
        add_data_source_frame = ttk.Frame(sources_frame)
        add_data_source_frame.pack(side="right")
        
        self.add_data_source_entry = ttk.Entry(add_data_source_frame, width=18)
        self.add_data_source_entry.pack(side="left", padx=(0, 5))
        self.add_data_source_entry.bind('<Return>', self.add_data_source_from_entry)
        self.add_data_source_entry.insert(0, "Enter data source...")
        self.add_data_source_entry.bind('<FocusIn>', lambda e: self.add_data_source_entry.delete(0, tk.END) if self.add_data_source_entry.get() == "Enter data source..." else None)
        
        add_data_source_btn = ttk.Button(add_data_source_frame, text="+ Add Data Source", 
                                      command=self.add_data_source_from_entry, width=18)
        add_data_source_btn.pack(side="left")
        
        # Create the main data sources notebook
        self.data_sources_notebook = ttk.Notebook(parent)
        self.data_sources_notebook.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # Create tabs for each data source
        for data_source in self.get_all_data_sources():
            config = self.data_sources[data_source]
            name = data_source
            
            # Create main tab for this data source
            data_source_tab = ttk.Frame(self.data_sources_notebook)
            self.data_sources_notebook.add(data_source_tab, text=data_source)
            
            # Create the data source tab content
            self.create_data_source_tab_content(data_source_tab, data_source, config)
        
        # Update all tab indicators after creating tabs
        self.update_all_tab_indicators()
        
        # Update all tab colors based on coordinate maps
        self.update_all_tab_colors()
        
        # Bind tab change event to update primary data source
        self.data_sources_notebook.bind("<<NotebookTabChanged>>", self.on_tab_changed)
        
        # Bind right-click event for context menu
        self.data_sources_notebook.bind("<Button-3>", self.show_tab_context_menu)
        
        print("DEBUG: create_data_sources_notebook completed!")
    
    def create_destinations_notebook(self, parent):
        """Create a notebook with tabs for each destination"""
        print("DEBUG: Starting create_destinations_notebook...")
        
        # Add a label and button for the destinations section
        destinations_frame = ttk.Frame(parent)
        destinations_frame.pack(fill="x", padx=5, pady=(5,0))
        
        destinations_label = tk.Label(destinations_frame, text="DESTINATIONS", font=("Arial", 10, "bold"), fg="blue")
        destinations_label.pack(side="left")
        
        # Add text entry and button to create new destination
        add_destination_frame = ttk.Frame(destinations_frame)
        add_destination_frame.pack(side="right")
        
        self.add_destination_entry = ttk.Entry(add_destination_frame, width=18)
        self.add_destination_entry.pack(side="left", padx=(0, 5))
        self.add_destination_entry.bind('<Return>', self.add_destination_from_entry)
        self.add_destination_entry.insert(0, "Enter destination...")
        self.add_destination_entry.bind('<FocusIn>', lambda e: self.add_destination_entry.delete(0, tk.END) if self.add_destination_entry.get() == "Enter destination..." else None)
        
        add_destination_btn = ttk.Button(add_destination_frame, text="+ Add Destination", 
                                      command=self.add_destination_from_entry, width=18)
        add_destination_btn.pack(side="left")
        
        # Create the main destinations notebook
        self.destinations_notebook = ttk.Notebook(parent)
        # Bind to tab change event to update button references
        self.destinations_notebook.bind("<<NotebookTabChanged>>", self.on_destination_tab_changed)
        self.destinations_notebook.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # Create tabs for each destination
        for destination in self.get_all_destinations():
            config = self.destinations[destination]
            
            # Create main tab for this destination
            destination_tab = ttk.Frame(self.destinations_notebook)
            self.destinations_notebook.add(destination_tab, text=destination)
            
            # Create the destination tab content
            self.create_destination_tab_content(destination_tab, destination, config)
        
        print("DEBUG: create_destinations_notebook completed!")
    
    def clear_search_placeholder(self):
        """Clear the placeholder text when user focuses on search entry"""
        if self.data_source_search_entry.get() == "Search keys/values...":
            self.data_source_search_entry.delete(0, tk.END)
            self.data_source_search_entry.config(foreground='black')
    
    def restore_search_placeholder(self):
        """Restore the placeholder text when user leaves search entry empty"""
        if not self.data_source_search_entry.get().strip():
            self.data_source_search_entry.insert(0, "Search keys/values...")
            self.data_source_search_entry.config(foreground='gray')
            # Clear any search results
            self.clear_data_source_search()
    
    def search_data_sources(self, event=None):
        """Search through all data sources for the given substring in keys and values"""
        search_text = self.data_source_search_entry.get().strip()
        
        # Don't search if it's the placeholder text or empty
        if not search_text or search_text == "Search keys/values...":
            # If search is empty, close the results window
            if hasattr(self, 'search_results_window') and self.search_results_window.winfo_exists():
                self.search_results_window.destroy()
            return
        
        search_text_lower = search_text.lower()
        
        # Collect all matches across all data sources
        all_matches = []
        
        for data_source in self.get_all_data_sources():
            data = self.data_sources[data_source].get('data', {})
            if not data:
                continue
            
            # Search through the data
            for key, value_dict in data.items():
                # Skip if key is None or not a valid type
                if key is None:
                    continue
                
                # Check if search term matches the key
                try:
                    key_str = str(key).lower()
                    if search_text_lower in key_str:
                        all_matches.append({
                            'data_source': data_source,
                            'key': key,
                            'match_type': 'key',
                            'match_value': str(key)
                        })
                except (AttributeError, TypeError):
                    # Skip keys that can't be converted to string
                    continue
                
                # Check if search term matches any value in the nested dictionary
                if isinstance(value_dict, dict):
                    for field, value in value_dict.items():
                        try:
                            value_str = str(value).lower()
                            if search_text_lower in value_str:
                                all_matches.append({
                                    'data_source': data_source,
                                    'key': key,
                                    'match_type': 'value',
                                    'field': field,
                                    'match_value': str(value)
                                })
                        except (AttributeError, TypeError):
                            # Skip values that can't be converted to string
                            continue
        
        # Update or create the results window
        self.update_search_results(search_text, all_matches)
    
    def update_search_results(self, search_text, matches):
        """Update or create search results window"""
        # Check if window exists and is valid
        window_exists = hasattr(self, 'search_results_window') and self.search_results_window.winfo_exists()
        
        if not window_exists:
            # Create the window for the first time
           self.create_search_results_window()
        
        center_window_over_parent(self.search_results_window)
        # Update the window title
        self.search_results_window.title(f"Search Results: '{search_text}'")
        
        # Update the count label
        self.search_count_label.config(text=f"Found {len(matches)} match(es) for '{search_text}'")
        
        # Clear existing items in the treeview
        for item in self.search_tree.get_children():
            self.search_tree.delete(item)
        
        # Add new matches to treeview
        for match in matches:
            data_source = match['data_source']
            key = match['key']
            match_type = match['match_type']
            field = match.get('field', '-')
            value = match.get('match_value', '-')
            
            # Truncate long values
            if len(value) > 100:
                value = value[:97] + "..."
            
            self.search_tree.insert("", "end", values=(data_source, key, match_type, field, value))
        
        # If no matches, show a message in the treeview
        if not matches:
            self.search_tree.insert("", "end", values=("No matches found", "", "", "", ""))
    
    def create_search_results_window(self):
        """Create the search results window (called once)"""
        # Create results window
        self.search_results_window = tk.Toplevel(self.root)
        self.search_results_window.title("Search Results")
        self.search_results_window.geometry("800x500")
        
        # Make window stay on top initially, but allow user to move it
        self.search_results_window.transient(self.root)
        
        # Create frame with scrollbar
        main_frame = ttk.Frame(self.search_results_window)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Add label showing count (will be updated)
        self.search_count_label = tk.Label(main_frame, 
                                           text="Searching...",
                                           font=("Arial", 10, "bold"))
        self.search_count_label.pack(anchor="w", pady=(0, 5))
        
        # Create treeview for results
        tree_frame = ttk.Frame(main_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        
        # Scrollbars
        vsb = ttk.Scrollbar(tree_frame, orient="vertical")
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal")
        
        # Treeview (store as instance variable)
        self.search_tree = ttk.Treeview(tree_frame, 
                                       columns=("Data Source", "Key", "Match Type", "Field", "Value"),
                                       show="headings",
                                       yscrollcommand=vsb.set,
                                       xscrollcommand=hsb.set)
        
        vsb.config(command=self.search_tree.yview)
        hsb.config(command=self.search_tree.xview)
        
        # Configure columns
        self.search_tree.heading("Data Source", text="Data Source")
        self.search_tree.heading("Key", text="Key")
        self.search_tree.heading("Match Type", text="Match Type")
        self.search_tree.heading("Field", text="Field")
        self.search_tree.heading("Value", text="Value")
        
        self.search_tree.column("Data Source", width=120)
        self.search_tree.column("Key", width=150)
        self.search_tree.column("Match Type", width=100)
        self.search_tree.column("Field", width=120)
        self.search_tree.column("Value", width=280)
        
        # Pack treeview and scrollbars
        self.search_tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)
        
        # Bind double-click to open data viewer
        self.search_tree.bind("<Double-Button-1>", self.on_search_result_double_click)
        
        # Add close button
        close_btn = ttk.Button(main_frame, text="Close", command=self.search_results_window.destroy)
        close_btn.pack(pady=(10, 0))
    
    def on_search_result_double_click(self, event):
        """Handle double-click on search result to open data viewer"""
        # Get the selected item
        selection = self.search_tree.selection()
        if not selection:
            return
        
        # Get the item's values
        item = selection[0]
        values = self.search_tree.item(item, 'values')
        
        # Extract the data source (first column)
        if values and len(values) > 0:
            data_source = values[0]
            
            # Check if it's a valid data source (not the "No matches found" row)
            if data_source and data_source != "No matches found" and data_source in self.data_sources:
                # Get the current search text to pass to the viewer
                search_text = self.data_source_search_entry.get().strip()
                if search_text == "Search keys/values...":
                    search_text = ""
                
                # Open the data viewer for this data source with the search term
                self.view_data(data_source=data_source, initial_search=search_text)
    
    def clear_data_source_search(self):
        """Clear the search entry and close any open search results"""
        self.data_source_search_entry.delete(0, tk.END)
        self.data_source_search_entry.insert(0, "Search keys/values...")
        self.data_source_search_entry.config(foreground='gray')
        
        # Close search results window if it exists
        if hasattr(self, 'search_results_window') and self.search_results_window.winfo_exists():
            self.search_results_window.destroy()
    
    def update_all_tab_indicators(self):
        """Update indicators for all data source tabs"""
        for data_source in self.get_all_data_sources():
            self.add_checkmark_to_tab(data_source)
    
    def show_tab_context_menu(self, event):
        """Show context menu when right-clicking on a tab"""
        # Get the tab index at the click position
        tab_index = self.data_sources_notebook.index(f"@{event.x},{event.y}")
        if tab_index is None:
            return
        
        # Get the data source from the tab text (remove check mark if present)
        tab_text = self.data_sources_notebook.tab(tab_index, "text")
        data_source = self.strip_tab_indicators(tab_text)  # Remove indicators to get base data source
        
        # Create context menu
        context_menu = tk.Menu(self.root, tearoff=0)
        context_menu.add_command(label="Rename data source", 
                               command=lambda: self.rename_data_source(data_source))
        context_menu.add_separator()
        context_menu.add_command(label="Delete data source", 
                               command=lambda: self.delete_data_source(data_source))
        
        # Show the context menu
        try:
            context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            context_menu.grab_release()
    
    def rename_data_source(self, old_data_source):
        """Rename a data source"""
        # Ask for new name
        new_data_source = tk.simpledialog.askstring("Rename data source", 
                                                 f"Enter new name for '{old_data_source}':",
                                                 parent=self.root,
                                                 initialvalue=old_data_source)
        
        if not new_data_source or new_data_source.strip() == "":
            return
        
        new_data_source = new_data_source.strip()
        
        # Check if new name already exists
        if new_data_source in self.data_sources:
            messagebox.showerror("Error", f"data source '{new_data_source}' already exists")
            return
        
        # Check if it's the same name
        if new_data_source == old_data_source:
            return
        
        # Get the old configuration
        old_config = self.data_sources[old_data_source].copy()
        
        # Create new data source with the new name
        self.data_sources[new_data_source] = old_config
        
        # Path is now stored in config, so no need to update instance attributes
        print(f"DEBUG: Renamed data source from {old_data_source} to {new_data_source}")
        
        # Update UI entries dictionary if it exists
        if hasattr(self, 'data_source_ui_entries') and old_data_source in self.data_source_ui_entries:
            self.data_source_ui_entries[new_data_source] = self.data_source_ui_entries[old_data_source]
            del self.data_source_ui_entries[old_data_source]
        
        # Update file entries dictionary if it exists
        if hasattr(self, 'data_source_file_entries') and old_data_source in self.data_source_file_entries:
            self.data_source_file_entries[new_data_source] = self.data_source_file_entries[old_data_source]
            del self.data_source_file_entries[old_data_source]
        
        # Update widget references dictionary
        if old_data_source in self.data_source_widgets:
            self.data_source_widgets[new_data_source] = self.data_source_widgets[old_data_source]
            del self.data_source_widgets[old_data_source]
        
        # Remove old data source
        del self.data_sources[old_data_source]
        
        # Find and recreate the tab to update all bindings
        if hasattr(self, 'data_sources_notebook'):
            for i in range(self.data_sources_notebook.index("end")):
                tab_text = self.data_sources_notebook.tab(i, "text")
                base_tab_text = tab_text.replace(" ✓", "")
                
                if base_tab_text == old_data_source:
                    # Get the tab widget
                    tab_widget = self.data_sources_notebook.nametowidget(self.data_sources_notebook.tabs()[i])
                    
                    # Destroy all children of the tab to clear it
                    for child in tab_widget.winfo_children():
                        child.destroy()
                    
                    # Recreate the full tab content with the new data source name
                    self.create_data_source_tab_content(tab_widget, new_data_source, self.data_sources[new_data_source])
                    
                    # Update the tab text using centralized function
                    new_tab_text = self.get_tab_text_with_indicators(new_data_source)
                    self.data_sources_notebook.tab(i, text=new_tab_text)
                    print(f"Recreated tab content for renamed data source: {new_data_source}")
                    break
        
        messagebox.showinfo("Success", f"data source renamed from '{old_data_source}' to '{new_data_source}'")
    
    def update_tab_text(self, old_data_source, new_data_source):
        """Update the tab text without destroying the tab content"""
        if not hasattr(self, 'data_sources_notebook'):
            return
        
        # Find the tab with the old data source name
        for i in range(self.data_sources_notebook.index("end")):
            tab_text = self.data_sources_notebook.tab(i, "text")
            # Remove indicators to get the base data source
            base_tab_text = self.strip_tab_indicators(tab_text)
            
            if base_tab_text == old_data_source:
                # Update the tab text using centralized function
                new_tab_text = self.get_tab_text_with_indicators(new_data_source)
                self.data_sources_notebook.tab(i, text=new_tab_text)
                print(f"Updated tab text from '{tab_text}' to '{new_tab_text}'")
                break
    
    def get_tab_text_with_indicators(self, data_source):
        """Centralized function to build tab text with appropriate indicators.
        Returns the formatted tab text based on data source state.
        
        Indicators:
        - 🗺️ = coordinate map exists for current destination
        - ✓ = data is populated
        """
        if data_source not in self.data_sources:
            return data_source
        
        base_text = data_source
        
        # Check if data is populated
        has_data = bool(self.data_sources[data_source].get('data', {}))
        
        # Check if coordinate map exists for current destination
        has_coordinate_map = False
        current_destination = self.get_current_destination()
        if current_destination:
            coord_vals = self.data_sources[data_source].get('coordinate_values', {})
            if current_destination in coord_vals:
                destination_coords = coord_vals[current_destination]
                has_coordinate_map = bool(destination_coords and len(destination_coords) > 0)
        
        # Build text with indicators (order: map indicator first, then data indicator)
        if has_coordinate_map:
            base_text += "🗺️"
        if has_data:
            base_text += "✓"
        
        return base_text
    
    def strip_tab_indicators(self, tab_text):
        """Centralized function to remove indicators from tab text.
        Returns the base data source name without indicators.
        
        Removes:
        - 🗺️ = coordinate map indicator
        - ✓ = data populated indicator
        """
        return tab_text.replace("🗺️", "").replace("✓", "")
    
    def add_checkmark_to_tab(self, data_source):
        """Add a checkmark to the tab to indicate data is populated"""
        if not hasattr(self, 'data_sources_notebook'):
            return
        
        # Use update_tab_color_for_data_source which handles both data and coordinate map indicators
        self.update_tab_color_for_data_source(data_source)
    
    def update_tab_color_for_data_source(self, data_source, tab_index=None):
        """Update tab text to include indicators based on data source state"""
        if not hasattr(self, 'data_sources_notebook'):
            return
        
        # Find tab index if not provided
        if tab_index is None:
            for i in range(self.data_sources_notebook.index("end")):
                tab_text = self.data_sources_notebook.tab(i, "text")
                base_data_source = self.strip_tab_indicators(tab_text)
                if base_data_source == data_source:
                    tab_index = i
                    break
        
        if tab_index is None:
            return
        
        # Get the properly formatted tab text with indicators
        try:
            new_text = self.get_tab_text_with_indicators(data_source)
            self.data_sources_notebook.tab(tab_index, text=new_text)
        except Exception as e:
            print(f"Warning: Could not update tab indicator for {data_source}: {e}")
    
    def update_all_tab_colors(self):
        """Update indicators for all data source tabs based on coordinate maps"""
        if not hasattr(self, 'data_sources_notebook'):
            return
        
        for i in range(self.data_sources_notebook.index("end")):
            tab_text = self.data_sources_notebook.tab(i, "text")
            base_data_source = self.strip_tab_indicators(tab_text)
            if base_data_source in self.data_sources:
                self.update_tab_color_for_data_source(base_data_source, i)
    
    def delete_data_source(self, data_source):
        """Delete a data source"""
        # Confirm deletion
        if not messagebox.askyesno("Confirm Delete", 
                                  f"Are you sure you want to delete data source '{data_source}'?\n\nThis action cannot be undone."):
            return
        
        # Remove from data sources
        if data_source in self.data_sources:
            del self.data_sources[data_source]
        
        # Remove from widget references
        if data_source in self.data_source_widgets:
            del self.data_source_widgets[data_source]
        
        # Path is now stored in config, so no need to clean up instance attributes

        # Clean up entries before refreshing
        self.cleanup_entries()
        
        # Refresh the interface
        self.refresh_data_sources_notebook()
        
        messagebox.showinfo("Success", f"data source '{data_source}' deleted")
    
    def refresh_data_sources_notebook(self):
        """Add new data source tabs and remove old ones to match current data sources"""
        if not hasattr(self, 'data_sources_notebook'):
            return
        
        # Get current data sources
        current_data_sources = set(self.get_all_data_sources())
        
        # Get all existing tab texts and their indices
        existing_tabs = {}
        for i in range(self.data_sources_notebook.index("end")):
            tab_text = self.data_sources_notebook.tab(i, "text")
            existing_tabs[tab_text] = i
        
        # Remove tabs that no longer exist in data sources
        tabs_to_remove = []
        for tab_text, tab_index in existing_tabs.items():
            if tab_text not in current_data_sources:
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
                print(f"Removed tab for data source: {tab_text}")
        
        # Clean up entries after removing tabs
        self.cleanup_entries()
        
        # Add tabs for any new data sources that don't exist yet
        for data_source in current_data_sources:
            if data_source not in existing_tabs:
                config = self.data_sources[data_source]
                name = data_source
                
                # Create main tab for this data source
                data_source_tab = ttk.Frame(self.data_sources_notebook)
                self.data_sources_notebook.add(data_source_tab, text=data_source)
                
                # Create the data source tab content
                self.create_data_source_tab_content(data_source_tab, data_source, config)
                
                print(f"Added new tab for data source: {data_source}")
        
        # Update all tab indicators
        self.update_all_tab_indicators()
        
        # Update all tab colors based on coordinate maps
        self.update_all_tab_colors()
        
        # Switch to the first tab if any exist
        if self.data_sources_notebook.index("end") > 0:
            self.data_sources_notebook.select(0)
    
    def refresh_destinations_notebook(self):
        """Refresh destinations notebook when destinations are added/removed"""
        if not hasattr(self, 'destinations_notebook'):
            return
        
        # Get current destinations
        current_destinations = set(self.get_all_destinations())
        
        # Get existing tabs
        existing_tabs = {}
        for i in range(self.destinations_notebook.index("end")):
            tab_text = self.destinations_notebook.tab(i, "text")
            existing_tabs[tab_text] = i
        
        # Find tabs to remove
        tabs_to_remove = []
        for tab_text, tab_index in existing_tabs.items():
            if tab_text not in current_destinations:
                tabs_to_remove.append(tab_index)
        
        # Remove tabs in reverse order to maintain indices
        for tab_index in sorted(tabs_to_remove, reverse=True):
            tab_text = None
            for text, idx in existing_tabs.items():
                if idx == tab_index:
                    tab_text = text
                    break
            self.destinations_notebook.forget(tab_index)
            if tab_text:
                print(f"Removed tab for destination: {tab_text}")
        
        # Add tabs for any new destinations that don't exist yet
        for destination in current_destinations:
            if destination not in existing_tabs:
                config = self.destinations[destination]
                
                # Create main tab for this destination
                destination_tab = ttk.Frame(self.destinations_notebook)
                self.destinations_notebook.add(destination_tab, text=destination)
                
                # Create the destination tab content
                self.create_destination_tab_content(destination_tab, destination, config)
                
                print(f"Added new tab for destination: {destination}")
        
        # Switch to the first tab if any exist
        if self.destinations_notebook.index("end") > 0:
            self.destinations_notebook.select(0)
    
    def create_data_source_tab_content(self, parent, data_source, config):
        """Create the content for a data source tab, including file selection and nested configuration tabs"""
        name = data_source
        
        # Create file selection frame at the top
        file_frame = ttk.LabelFrame(parent, text=f"File Selection")
        file_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # File selection row
        file_row = tk.Frame(file_frame)
        file_row.pack(fill=tk.X, padx=5, pady=5)
        
        # Label
        label = tk.Label(file_row, text=f"File:", width=20)
        label.pack(side=tk.LEFT, padx=5)
        
        # Entry
        entry = tk.Entry(file_row)
        entry._data_source = data_source  # Store data source name for repopulation
        # Initialize with current path from config
        current_path = config.get('path', '')
        if current_path:
            entry.insert(0, current_path)
        entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
        # Store reference to the file entry for this data source
        if not hasattr(self, 'data_source_file_entries'):
            self.data_source_file_entries = {}
        self.data_source_file_entries[data_source] = entry
        
        # Buttons frame
        buttons = tk.Frame(file_row)
        buttons.pack(side=tk.RIGHT)
        
        # Browse button
        tk.Button(buttons, text="Browse",
                  command=lambda e=entry, ds=data_source: self.browse_data_source(e, ds)).pack(side=tk.LEFT, padx=2)
        
        # Open button
        tk.Button(buttons, text="Open",
                  command=lambda e=entry: self.open_file(e)).pack(side=tk.LEFT, padx=2)
        
        # Configure button
        tk.Button(buttons, text="Configure",
                  command=lambda n=name: self.configure(n)).pack(side=tk.LEFT, padx=2)
        
        # Generate button
        tk.Button(buttons, text="Generate",
                  command=lambda dt=data_source: self.generate_data_source(dt)).pack(side=tk.LEFT, padx=2)
        
        # View button
        tk.Button(buttons, text="View",
                  command=lambda ds=data_source: self.view_data(data_source=ds)).pack(side=tk.LEFT, padx=2)
        
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
        self.create_coordinates_tab(coordinates_tab, data_source)
        self.create_filters_tab(filters_tab, data_source)
        self.create_transform_tab(transform_tab, data_source)
        
        # Store reference to the config notebook for this data source
        if data_source not in self.data_source_widgets:
            self.data_source_widgets[data_source] = {}
        self.data_source_widgets[data_source]['config_notebook'] = config_notebook
    
    def create_destination_tab_content(self, parent, destination, config):
        """Create the content for a destination tab"""
        # Create main content frame with left and right sections
        content_frame = tk.Frame(parent)
        content_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # Left section for configuration fields
        left_section = tk.Frame(content_frame)
        left_section.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(5, 10))
        
        # Right section for action buttons
        right_section = tk.Frame(content_frame, relief=tk.RAISED, borderwidth=1)
        right_section.pack(side=tk.RIGHT, fill=tk.Y, padx=(10, 5))
        
        # === LEFT SECTION: Configuration Fields ===
        
        # Datasheets Row
        ds_frame = tk.Frame(left_section)
        ds_frame.pack(fill=tk.X, pady=2)
        
        # Help button
        ds_help_btn = tk.Button(ds_frame, text="?", width=2, 
                                command=lambda: self.show_help("datasheets_destination.txt"))
        ds_help_btn.pack(side=tk.LEFT, padx=(5, 2))
        
        ds_label = tk.Label(ds_frame, text="Datasheets (Destination)", width=30)
        ds_label.pack(side=tk.LEFT, padx=(0, 5))
        
        ds_entry = tk.Entry(ds_frame)
        ds_entry._destination = destination
        ds_entry._variable_name = "datasheets"
        if config.get('path'):
            ds_entry.insert(0, config['path'])
        ds_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
        # Store reference to datasheet entry
        if destination not in self.destination_widgets:
            self.destination_widgets[destination] = {}
        self.destination_widgets[destination]['datasheet_entry'] = ds_entry
        
        ds_buttons = tk.Frame(ds_frame)
        ds_buttons.pack(side=tk.RIGHT)
        
        tk.Button(ds_buttons, text="Browse",
                  command=lambda: self.browse_datasheets(ds_entry, destination)).pack(side=tk.LEFT, padx=2)
        tk.Button(ds_buttons, text="Map Coordinates",
                  command=lambda: self.configure_ds(destination)).pack(side=tk.LEFT, padx=2)
        
        # Datasheet Coordinate Row
        coord_frame = tk.Frame(left_section)
        coord_frame.pack(fill=tk.X, pady=2)
        
        coord_help_btn = tk.Button(coord_frame, text="?", width=2, 
                                   command=lambda: self.show_help("datasheet_coordinate.txt"))
        coord_help_btn.pack(side=tk.LEFT, padx=(5, 2))
        
        coord_label = tk.Label(coord_frame, text="Datasheet Coordinate", width=30)
        coord_label.pack(side=tk.LEFT, padx=(0, 5))
        
        coord_entry = tk.Entry(coord_frame)
        coord_entry._destination = destination
        coord_entry._variable_name = "datasheet_coord"
        if config.get('datasheet_coord'):
            coord_entry.insert(0, config['datasheet_coord'])
        coord_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        self.destination_widgets[destination]['datasheet_coord_entry'] = coord_entry
        
        # Datasheet Prefix Row
        prefix_frame = tk.Frame(left_section)
        prefix_frame.pack(fill=tk.X, pady=2)
        
        prefix_help_btn = tk.Button(prefix_frame, text="?", width=2, 
                                    command=lambda: self.show_help("datasheet_prefix.txt"))
        prefix_help_btn.pack(side=tk.LEFT, padx=(5, 2))
        
        prefix_label = tk.Label(prefix_frame, text="Datasheet Prefix", width=30)
        prefix_label.pack(side=tk.LEFT, padx=(0, 5))
        
        prefix_entry = tk.Entry(prefix_frame)
        prefix_entry._destination = destination
        prefix_entry._variable_name = "ds_str"
        if config.get('ds_str'):
            prefix_entry.insert(0, config['ds_str'])
        prefix_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        self.destination_widgets[destination]['ds_str_entry'] = prefix_entry
        
        # Rows per Sheet Row
        rows_frame = tk.Frame(left_section)
        rows_frame.pack(fill=tk.X, pady=2)
        
        rows_help_btn = tk.Button(rows_frame, text="?", width=2, 
                                  command=lambda: self.show_help("rows_per_sheet.txt"))
        rows_help_btn.pack(side=tk.LEFT, padx=(5, 2))
        
        rows_label = tk.Label(rows_frame, text="Rows per Sheet", width=30)
        rows_label.pack(side=tk.LEFT, padx=(0, 5))
        
        rows_entry = tk.Entry(rows_frame)
        rows_entry._destination = destination
        rows_entry._variable_name = "rows_per_sheet"
        if config.get('rows_per_sheet'):
            rows_entry.insert(0, str(config['rows_per_sheet']))
        rows_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        self.destination_widgets[destination]['rows_per_sheet_entry'] = rows_entry
        
        # Significant Figures Row
        sig_figs_frame = tk.Frame(left_section)
        sig_figs_frame.pack(fill=tk.X, pady=2)
        
        sig_figs_help_btn = tk.Button(sig_figs_frame, text="?", width=2, 
                                      command=lambda: self.show_help("significant_figures.txt"))
        sig_figs_help_btn.pack(side=tk.LEFT, padx=(5, 2))
        
        sig_figs_label = tk.Label(sig_figs_frame, text="Significant Figures", width=30)
        sig_figs_label.pack(side=tk.LEFT, padx=(0, 5))
        
        sig_figs_entry = tk.Entry(sig_figs_frame)
        sig_figs_entry._destination = destination
        sig_figs_entry._variable_name = "sig_figs"
        if config.get('sig_figs'):
            sig_figs_entry.insert(0, str(config['sig_figs']))
        sig_figs_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        self.destination_widgets[destination]['sig_figs_entry'] = sig_figs_entry
        
        # Rounding Tolerance Row
        tolerance_frame = tk.Frame(left_section)
        tolerance_frame.pack(fill=tk.X, pady=2)
        
        tolerance_help_btn = tk.Button(tolerance_frame, text="?", width=2, 
                                       command=lambda: self.show_help("rounding_tolerance.txt"))
        tolerance_help_btn.pack(side=tk.LEFT, padx=(5, 2))
        
        tolerance_label = tk.Label(tolerance_frame, text="Rounding Tolerance", width=30)
        tolerance_label.pack(side=tk.LEFT, padx=(0, 5))
        
        tolerance_entry = tk.Entry(tolerance_frame)
        tolerance_entry._destination = destination
        tolerance_entry._variable_name = "rounding_tolerance"
        if config.get('rounding_tolerance'):
            tolerance_entry.insert(0, str(config['rounding_tolerance']))
        tolerance_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        self.destination_widgets[destination]['rounding_tolerance_entry'] = tolerance_entry
        
        # === RIGHT SECTION: Action Buttons ===
        
        # Add title for the right section
        action_title = tk.Label(right_section, text="ACTIONS", font=("Arial", 9, "bold"), fg="darkgreen")
        action_title.pack(pady=(10, 5))
        
        # Color coding method
        color_frame = tk.Frame(right_section)
        color_frame.pack(fill=tk.X, pady=5)
        
        # Help button
        color_help_btn = tk.Button(color_frame, text="?", width=2, 
                                   command=lambda: self.show_help("color_coding_method.txt"))
        color_help_btn.pack(side=tk.LEFT, padx=(5, 2))
        
        color_label = tk.Label(color_frame, text="Color Coding Method:", width=20)
        color_label.pack(side=tk.LEFT, padx=(0, 5))
        
        # Use instance variable if it exists, otherwise create it
        if not hasattr(self, 'color_coding_var'):
            self.color_coding_var = tk.StringVar(value="new_red_old_green")
        color_dropdown = ttk.Combobox(color_frame, textvariable=self.color_coding_var, 
                                     values=["None (Black)", "new_red_old_green", "new_red", "new_red_and_highlight"], 
                                     state="readonly", width=20)
        color_dropdown.pack(side=tk.LEFT, padx=5)
        
        # Fill mode option
        fill_mode_frame = tk.Frame(right_section)
        fill_mode_frame.pack(fill=tk.X, pady=5)
        
        fill_mode_label = tk.Label(fill_mode_frame, text="New Tags Sheet Fill Mode:", width=20)
        fill_mode_label.pack(side=tk.LEFT, padx=(0, 5))
        
        # Use instance variable if it exists, otherwise create it
        if not hasattr(self, 'fill_mode_var'):
            self.fill_mode_var = tk.StringVar(value="continue_last")
        fill_mode_dropdown = ttk.Combobox(fill_mode_frame, textvariable=self.fill_mode_var, 
                                         values=["fill_blanks", "continue_last", "always_new", "ignore"], 
                                         state="readonly", width=20)
        fill_mode_dropdown.pack(side=tk.LEFT, padx=5)
        
        # Update matched tags option
        update_matched_frame = tk.Frame(right_section)
        update_matched_frame.pack(fill=tk.X, pady=5)
        
        update_matched_label = tk.Label(update_matched_frame, text="Update Matched Tags:", width=20)
        update_matched_label.pack(side=tk.LEFT, padx=(0, 5))
        
        # Use instance variable if it exists, otherwise create it
        if not hasattr(self, 'update_matched_var'):
            self.update_matched_var = tk.StringVar(value="update")
        update_matched_dropdown = ttk.Combobox(update_matched_frame, textvariable=self.update_matched_var, 
                                               values=["update", "skip"], 
                                               state="readonly", width=20)
        update_matched_dropdown.pack(side=tk.LEFT, padx=5)
        
        # Green highlighting option
        green_highlight_frame = tk.Frame(right_section)
        green_highlight_frame.pack(fill=tk.X, pady=5)
        
        # Use instance variable if it exists, otherwise create it
        if not hasattr(self, 'disable_green_highlight_var'):
            self.disable_green_highlight_var = tk.BooleanVar(value=False)
        disable_green_highlight_checkbox = tk.Checkbutton(
            green_highlight_frame,
            text="Disable Green Cell Highlighting",
            variable=self.disable_green_highlight_var,
            font=("Arial", 9)
        )
        disable_green_highlight_checkbox.pack(anchor=tk.W)
        
        # Append suffix option for green highlighted cells
        if not hasattr(self, 'append_suffix_to_green_var'):
            self.append_suffix_to_green_var = tk.BooleanVar(value=False)
        append_suffix_checkbox = tk.Checkbutton(
            green_highlight_frame,
            text="Append suffix to green highlighted cells",
            variable=self.append_suffix_to_green_var,
            font=("Arial", 9)
        )
        append_suffix_checkbox.pack(anchor=tk.W)
        
        # Clear highlighting option for matched cells
        if not hasattr(self, 'clear_highlighting_on_match_var'):
            self.clear_highlighting_on_match_var = tk.BooleanVar(value=False)
        clear_highlighting_checkbox = tk.Checkbutton(
            green_highlight_frame,
            text="Clear highlighting from matched cells",
            variable=self.clear_highlighting_on_match_var,
            font=("Arial", 9)
        )
        clear_highlighting_checkbox.pack(anchor=tk.W)
        
        # Add/Update button (per destination)
        generate_button = tk.Button(right_section, text="Add/Update",
                  command=self.add_datasheets, font=("Arial", 10, "bold"), 
                  bg="green", fg="white", padx=20, pady=8, width=12)
        generate_button.pack(pady=5, padx=10)
        self.destination_widgets[destination]['generate_button'] = generate_button
        
        # Stop button (per destination)
        stop_button = tk.Button(right_section, text="Stop",
                  command=self.set_halt_flag, bg="red", fg="white", state="disabled",
                  font=("Arial", 10, "bold"), padx=20, pady=8, width=12)
        stop_button.pack(pady=5, padx=10)
        self.destination_widgets[destination]['stop_button'] = stop_button
        
        # Add status label (per destination)
        status_label = tk.Label(right_section, text="Ready", fg="black", font=("Arial", 9))
        status_label.pack(anchor=tk.W, padx=5, pady=(5,0))
        self.destination_widgets[destination]['status_label'] = status_label
        
        # Also set as the current buttons if this is the first destination or if we need a default
        # This ensures backward compatibility with code that references self.generate_button directly
        if not hasattr(self, 'generate_button') or self.generate_button is None:
            self.generate_button = generate_button
            self.stop_button = stop_button
            self.status_label = status_label
        else:
            # Update to the most recently created destination's buttons
            self.generate_button = generate_button
            self.stop_button = stop_button
            self.status_label = status_label
    
    def generate_data_source(self, data_source):
        """Generic method to generate data for any data source"""
        # Use the generic approach for all data sources - truly extensible
        self.generate_generic_data(data_source)
        self.refresh_tab_content()
    
    def generate_generic_data(self, data_source):
        """Generate data for any data source using the centralized system"""

        config = self.data_sources[data_source]
        name = data_source
        
        print(f"DEBUG: generate_generic_data for {name}")
        
        # Get the file path from the entry widget first, fall back to config
        file_path = None
        if hasattr(self, 'data_source_file_entries') and data_source in self.data_source_file_entries:
            entry = self.data_source_file_entries[data_source]
            file_path = entry.get().strip()
            print(f"DEBUG: Retrieved file_path from entry: {file_path}")
        
        # Fall back to config if entry is empty or doesn't exist
        if not file_path:
            file_path = self.data_sources[data_source]['path']
            print(f"DEBUG: Retrieved file_path from config: {file_path}")
        
        if not file_path:
            messagebox.showwarning("Warning", f"No file path set for {name}")
            return
        
        # Determine file type and process accordingly
        file_extension = os.path.splitext(file_path)[1].lower()
        
        if file_extension == '.json':
            # Process JSON file
            try:
                data = load_dict_from_json(file_path)
                self.data_sources[data_source]['data'] = data
                # Update the tab indicator to show data status
                self.add_checkmark_to_tab(data_source)
                # Update coordinates section combo box with new data
                self.update_coordinates_combo_box(data_source)
                messagebox.showinfo("Success", f"Successfully loaded {name} data from JSON file.")
                print(f"Loaded {name} data from JSON: {file_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load JSON file:\n{str(e)}")
                
        elif file_extension in ['.xlsx', '.xls', '.xlsm']:
            # Show custom dialog for Excel processing
            choice = self.show_excel_processing_dialog(data_source, file_path)
            
            if choice == "datasheet":
                # Process as datasheet
                self.load_data_source_from_datasheet(data_source, file_path)
            elif choice == "index":
                # Process as index (generate dictionary)
                self.process_excel_as_index(file_path, data_source)
            # If choice is None (Cancel), do nothing
        else:
            messagebox.showerror("Error", f"Unsupported file type for {name}: {file_extension}")
    
    def process_excel_as_index(self, file_path, data_source):
        """Process Excel file as index (generate dictionary from headers)"""
        
        config = self.data_sources[data_source]
        name = data_source
        
        # Get headers and selected sheets
        headers = self.data_sources[data_source]['headers']
        selected_sheets = self.data_sources[data_source]['selected_sheets']
        
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
                self.data_sources[data_source]['headers'] = result["headers"]
                self.data_sources[data_source]['selected_sheets'] = result["selected_sheets"]
                self.blank_cell_tolerance = result["tolerance"]
                
                # Use the updated configuration
                headers = result["headers"]
                selected_sheets = result["selected_sheets"]
                print(f"{name} configuration updated.")
            else:
                print(f"{name} configuration cancelled.")
                return
        
        # Generate data using the centralized data source pattern
        data = generate_dictionary_from_xlsx(file_path, headers,
                                          parent=self.root, selected_sheets=selected_sheets,
                                          max_empty_allowed=self.blank_cell_tolerance)
        
        if data is not None:
            self.data_sources[data_source]['data'] = data
            # Update the tab indicator to show data status
            self.add_checkmark_to_tab(data_source)
            # Update coordinates section combo box with new data
            self.update_coordinates_combo_box(data_source)
            show_nested_dict_analysis(data)
            print(f"Generated {name}")
            self.refresh_tab_content()
        else:
            print(f"{name} generation cancelled or failed.")
    
    def refresh_data_source_frames(self):
        """Refresh all data source frames when new data sources are added"""
        # Remove existing data source frames
        for data_source in self.get_all_data_sources():
            # Get frame reference from data_source_widgets
            frame = self.data_source_widgets.get(data_source, {}).get('frame')
            if frame:
                frame.destroy()
                # Remove from data_source_widgets
                if data_source in self.data_source_widgets:
                    del self.data_source_widgets[data_source]['frame']
        
        # Recreate frames
        main_frame = self.root.winfo_children()[0]  # Get the main frame
        self.create_data_source_frames(main_frame)

        #self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def show_help(self, help_file):
        """Display help text from a file in a popup window"""
        help_path = os.path.join("Help", help_file)
        
        try:
            with open(help_path, 'r', encoding='utf-8') as f:
                help_text = f.read()
        except FileNotFoundError:
            help_text = f"Help file not found: {help_file}"
        except Exception as e:
            help_text = f"Error loading help: {str(e)}"
        
        # Create help window
        help_window = tk.Toplevel(self.root)
        help_window.title("Help")
        help_window.geometry("600x400")
        help_window.transient(self.root)
        
        # Center over the main window
        center_window_over_parent(help_window)
        
        # Create scrolled text widget
        text_widget = scrolledtext.ScrolledText(help_window, wrap=tk.WORD, font=("Consolas", 10))
        text_widget.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Insert help text
        text_widget.insert(1.0, help_text)
        text_widget.configure(state='disabled')
        
        # Add close button
        close_button = tk.Button(help_window, text="Close", command=help_window.destroy, width=10)
        close_button.pack(pady=(0, 10))

    def create_widgets(self):
        print("DEBUG: Starting create_widgets...")
        # Create menu bar
        print("DEBUG: Creating menu bar...")
        self.menu_bar = tk.Menu(self.root)
        self.root.config(menu=self.menu_bar)

        # Create File menu
        print("DEBUG: Creating File menu...")
        self.file_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label="File", menu=self.file_menu)
        
        # Add File menu items
        self.file_menu.add_command(label="New Project", command=self.new_project)
        self.file_menu.add_separator()
        self.file_menu.add_command(label="Load Settings", command=self.load_settings)
        self.file_menu.add_command(label="Save", command=self.save_settings, accelerator="Ctrl+S")
        self.file_menu.add_command(label="Save As", command=self.save_settings_as, accelerator="Ctrl+Shift+S")
        
        # Bind keyboard shortcuts
        self.root.bind('<Control-s>', lambda e: self.save_settings())
        self.root.bind('<Control-Shift-S>', lambda e: self.save_settings_as())

        # Create Commands menu
        print("DEBUG: Creating Commands menu...")
        self.command_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label="Commands", menu=self.command_menu)

        # Add menu items (Load/Save Settings removed - now in File menu)
        print("DEBUG: Setting up menu commands...")
        menu_commands = [
            ("Run xlsx search app", self.open_excel_search_app),
            ("Populate Headers on Datasheets", self.open_edit_xlsx),
            ("View Coordinate Value Data", self.display_coordinate_values),
            ("Delete newly added datasheets", self.delete_added_sheets),
            ("Rebuild tabs", self.refresh_tab_content),
            ("Sort Tabs", self.sort_tabs),
            ("Delete Certain Sheets by Prefix", self.delete_sheets_by_prefix),
            ("Excel Macros", self.open_excel_macros_window),
            ("Excel Regex Search App", self.open_excel_regex_search_app),
            ("Semantic Matcher", self.open_semantic_matcher),
            ("Release Excel", self.release_excel_connection),
            ("Release All Excel", self.release_all_excel_connections),
            ("Stop Datasheet Generation", self.set_halt_flag),
            ("refresh tab content", self.refresh_tab_content(force_rebuild=True)),
            ('update combo box', self.update_combo_boxes),
            ('update coordinates combo', self.update_all_coordinates_combo_boxes),
            ("Migrate to Multi-Destination Format", self.migrate_to_multi_destination),
        ]
        
        # Add dynamic menu items for each data source
        print("DEBUG: Adding dynamic menu items for data sources...")
        for data_source in self.get_all_data_sources():
            config = self.data_sources[data_source]
            name = data_source
            
            # data source specific menu items removed - now handled by browse button

        print("DEBUG: Adding menu commands to menu...")
        for label, command in menu_commands:
            self.command_menu.add_command(label=label, command=command)

        # Create main container frame
        print("DEBUG: Creating main container frame...")
        main_frame = tk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        self.entries = []  # Store entries for later reference

        # Create data sources notebook with tabs for each data source
        print("DEBUG: Creating data sources notebook...")
        self.create_data_sources_notebook(main_frame)

        # Add a visual separator between data sources and destinations
        print("DEBUG: Creating separator and destination container...")
        separator_frame = tk.Frame(main_frame, height=2, bg='gray')
        separator_frame.pack(fill=tk.X, pady=10)
        separator_frame.pack_propagate(False)

        # Create destinations notebook with tabs for each destination
        print("DEBUG: Creating destinations notebook...")
        self.create_destinations_notebook(main_frame)
        
        # Initialize global option variables (shared across all destinations)
        if not hasattr(self, 'color_coding_var'):
            self.color_coding_var = tk.StringVar(value="new_red_old_green")
        if not hasattr(self, 'fill_mode_var'):
            self.fill_mode_var = tk.StringVar(value="continue_last")
        if not hasattr(self, 'update_matched_var'):
            self.update_matched_var = tk.StringVar(value="update")
        if not hasattr(self, 'disable_green_highlight_var'):
            self.disable_green_highlight_var = tk.BooleanVar(value=False)
        
        # Initialize button references (will be set when first destination tab is created)
        self.generate_button = None
        self.stop_button = None
        self.status_label = None

        # Repopulate the entries list after all widgets are created
        self.repopulate_entries_list()

        print("DEBUG: create_widgets completed!")

    # endregion
    # region Tab Creation

    def init_excel(self):
        """Initialize Excel only when needed. 
        Now supports multiple open workbooks - will switch to the appropriate one or open it if needed."""
        if self.destination_datasheet:
            print("DEBUG: About to open workbook with ExcelManager...")
            # open_workbook now handles checking if already open and switching to it
            self.excel_mgr.open_workbook(self.destination_datasheet)
            print("DEBUG: Workbook opened/switched successfully with ExcelManager")

    def create_coordinates_tab(self, tab, data_source=None, sheet_names=None, destination=None):
        print(f"DEBUG: Starting create_coordinates_tab for data_source: {data_source}, destination: {destination}...")
        
        # Get current destination if not provided
        if destination is None:
            destination = self.get_current_destination()
            if destination is None:
                # Create default destination if none exist
                default_dest = "Default"
                if default_dest not in self.destinations:
                    self.add_destination(default_dest, {
                        'path': '', 'datasheet_coord': '', 'ds_str': '',
                        'rows_per_sheet': 1, 'sig_figs': 4, 'rounding_tolerance': 1e-2
                    })
                destination = default_dest

        def get_combo_values():
            """Get combo values for the specific data source"""
            if data_source:
                data = self.data_sources[data_source]['data']
                values = []
                if data:
                    # Get keys from the first entry
                    first_entry = list(data.values())[0]
                    if isinstance(first_entry, dict):
                        values = list(first_entry.keys())
                return {data_source: values}
            return {}

        def add_coordinate(data_source, entry, combo, listbox, dest):
            coord = entry.get()
            value = combo.get()
            if coord and value:
                # Access coordinate_values per destination
                if dest not in self.data_sources[data_source]['coordinate_values']:
                    self.data_sources[data_source]['coordinate_values'][dest] = {}
                coordinate_values = self.data_sources[data_source]['coordinate_values'][dest]
                coordinate_values[coord] = value
                self.data_sources[data_source]['coordinate_values'][dest] = coordinate_values
                
                # Only update the specific listbox that was changed, not all listboxes
                update_listbox(data_source, listbox, dest)
                
                # Update tab color since coordinate map changed
                self.update_tab_color_for_data_source(data_source)
                
                # Keep the entry text so user can add multiple coordinates easily

        def remove_coordinate(data_source, listbox, dest):
            selected_indices = listbox.curselection()
            if not selected_indices:
                return
            # Access coordinate_values per destination
            if dest not in self.data_sources[data_source]['coordinate_values']:
                self.data_sources[data_source]['coordinate_values'][dest] = {}
            coordinate_values = self.data_sources[data_source]['coordinate_values'][dest]
            for idx in reversed(selected_indices):
                coord = listbox.get(idx).split(':')[0].strip()
                if coord in coordinate_values:
                    del coordinate_values[coord]
            self.data_sources[data_source]['coordinate_values'][dest] = coordinate_values
            # Only update the specific listbox that was changed
            update_listbox(data_source, listbox, dest)
            
            # Update tab color since coordinate map changed
            self.update_tab_color_for_data_source(data_source)

        def clear_coordinates(data_source, listbox, dest):
            # Access coordinate_values per destination
            if dest not in self.data_sources[data_source]['coordinate_values']:
                self.data_sources[data_source]['coordinate_values'][dest] = {}
            self.data_sources[data_source]['coordinate_values'][dest] = {}
            # Only update the specific listbox that was changed
            update_listbox(data_source, listbox, dest)
            
            # Update tab color since coordinate map changed
            self.update_tab_color_for_data_source(data_source)

        def update_coordinate_display(full_selection, current_selection):
            """Callback function for centralized Excel selection monitoring - only updates coordinate entry"""
            try:
                # Update coordinate entry only
                entry_var.set(current_selection)
            except Exception as e:
                print(f"Error updating coordinate display: {e}")
        
        # Register this tab's callback with the centralized monitoring
        self.register_coordinate_callback(update_coordinate_display)
        
        # Store callback reference for cleanup
        tab._coordinate_callback = update_coordinate_display

        def update_listbox(data_source, listbox, dest):
            """Update a specific data source's listbox using centralized system"""
            try:
                # Check if the widget still exists
                if not listbox.winfo_exists():
                    return
                listbox.delete(0, tk.END)
                # Access coordinate_values per destination
                if dest not in self.data_sources[data_source]['coordinate_values']:
                    self.data_sources[data_source]['coordinate_values'][dest] = {}
                coordinate_values = self.data_sources[data_source]['coordinate_values'][dest]
                for key, value in coordinate_values.items():
                    # Add placeholder for conversion details
                    coord_display = f"{key}: {value}"
                    if key in self.data_sources[data_source]['coordinate_conversions']:
                        conv = self.data_sources[data_source]['coordinate_conversions'][key]
                        coord_display += f" [{conv.get('in_unit', '?')}->{conv.get('out_unit', '?')}]"
                    if key in self.data_sources[data_source]['coordinate_combinations']:
                        combo = self.data_sources[data_source]['coordinate_combinations'][key]
                        coord_display += f" [Combines: {', '.join(combo.get('combines', []))} ({combo.get('operation', 'add')})]"
                    listbox.insert(tk.END, coord_display)
            except tk.TclError:
                # Widget was destroyed, skip updating
                print(f"Warning: Could not update {data_source} listbox - widget may have been destroyed")
                return

        def update_listboxes():
            """Update all listboxes for all data sources using current destination"""
            current_dest = self.get_current_destination()
            if not current_dest:
                current_dest = destination  # Fallback to original destination
            for data_source in self.get_all_data_sources():
                # Get listbox reference from data_source_widgets
                listbox = self.data_source_widgets.get(data_source, {}).get('listbox')
                if listbox:
                    try:
                        # Check if the widget still exists before updating
                        if listbox.winfo_exists():
                            update_listbox(data_source, listbox, current_dest)
                    except tk.TclError:
                        # Widget was destroyed, skip updating
                        print(f"Warning: Could not update listbox for {data_source} - widget may have been destroyed")
                        continue

        def reinitialize():
            #init_excel()
            combo_values = get_combo_values()
            for data_source in self.get_all_data_sources():
                # Get combo reference from data_source_widgets
                combo = self.data_source_widgets.get(data_source, {}).get('combo')
                if combo:
                    combo['values'] = combo_values.get(data_source, [])
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
        right_frame = ttk.Frame(main_container, padding="5")
        right_frame.pack(side="right", fill="y", padx=(5, 0))
        right_frame.configure(width=300)  # Fixed width for right panel
        
        # UI Setup with Top Tag entry for the specific data source
        top_frame = ttk.Frame(left_frame)
        top_frame.pack(fill="x", pady=(0, 5))

        # Top Tag entry (only show for the specific data source)
        if data_source:
            current_top_tag = self.data_sources[data_source]['top_tag']
            config = self.data_sources[data_source]
            name = data_source
            print(f"DEBUG: Getting top_tag for {data_source}: '{current_top_tag}', config: {config}")
            
            top_tag_frame = ttk.Frame(top_frame)
            top_tag_frame.pack(fill="x", pady=(0, 5))
            
            ttk.Label(top_tag_frame, text=f"Top Tag for {name}:").pack(side="left")
            top_tag_entry = ttk.Entry(top_tag_frame, width=10)
            top_tag_entry.pack(side="left", padx=(5, 0))
            # Ensure current_top_tag is a string
            top_tag_value = str(current_top_tag) if current_top_tag is not None else ""
            top_tag_entry.insert(0, top_tag_value)
            
            # Store reference to the top tag entry for this data source
            if not hasattr(self, 'data_source_ui_entries'):
                self.data_source_ui_entries = {}
            if data_source not in self.data_source_ui_entries:
                self.data_source_ui_entries[data_source] = {}
            self.data_source_ui_entries[data_source]['top_tag_entry'] = top_tag_entry
            
            def update_top_tag():
                new_top_tag = top_tag_entry.get().strip()
                if new_top_tag:
                    self.data_sources[data_source]['top_tag'] = new_top_tag
                    print(f"Updated {name} top tag to: {new_top_tag}")
            
            ttk.Button(top_tag_frame, text="Update", command=update_top_tag).pack(side="left", padx=(5, 0))
            
            # Add help button for top tag
            help_btn = ttk.Button(top_tag_frame, text="?", width=3, 
                                 command=lambda: self.show_help("top_tag.txt"))
            help_btn.pack(side="left", padx=(5, 0))
            
            # Partial match checkbox (only show for the specific data source)
            partial_match_frame = ttk.Frame(top_frame)
            partial_match_frame.pack(fill="x", pady=(0, 5))
            
            current_partial_match = self.data_sources[data_source].get('partial_match', False)
            partial_match_var = tk.BooleanVar(value=current_partial_match)
            partial_match_checkbox = tk.Checkbutton(partial_match_frame, 
                                                   text="Use partial matching (find tag in cell text)", 
                                                   variable=partial_match_var)
            partial_match_checkbox.pack(side="left", padx=5)
            
            # Store reference to the partial match variable for this data source
            self.data_source_ui_entries[data_source]['partial_match_var'] = partial_match_var
            
            def update_partial_match():
                self.data_sources[data_source]['partial_match'] = partial_match_var.get()
                print(f"Updated {name} partial_match to: {partial_match_var.get()}")
            
            # Bind checkbox to update function
            partial_match_checkbox.config(command=update_partial_match)

        # Coordinate entry row with cell values and selection info
        coord_frame = ttk.Frame(top_frame)
        coord_frame.configure(width=50)
        coord_frame.pack(pady=(0, 5), anchor="w")
        
        coord_label = ttk.Label(coord_frame, text="Enter Key Coordinate:")
        coord_label.pack(side="left")

        entry_var = tk.StringVar()
        coord_entry = ttk.Entry(coord_frame, textvariable=entry_var)
        coord_entry.pack(side="left", fill="x", expand=True, padx=(5, 0))
        
        # Cell values display on same line
        cell_values_var = tk.StringVar()
        cell_values_var.set("Above: — | Left: —")
        cell_values_label = ttk.Label(coord_frame, textvariable=cell_values_var, 
                                     font=("Arial", 9), foreground="gray", width=40)
        cell_values_label.pack(side="left", padx=(10, 0))

        # Region selection display on same line
        region_var = tk.StringVar()
        region_var.set("Region: Single Cell")
        region_label = ttk.Label(coord_frame, textvariable=region_var, 
                                font=("Arial", 9), foreground="blue", width=20)
        region_label.pack(side="left", padx=(10, 0))
        
        # Refresh button for cell context (command will be configured after update_cell_context_display is defined)
        refresh_context_button = ttk.Button(coord_frame, text="🔄", width=3)
        refresh_context_button.pack(side="left", padx=(5, 0))

        # Define the update function for cell context (after variables are created)
        def update_cell_context_display():
            """Manually update region and cell values (Above/Left) display"""
            try:
                # Get current selection from Excel
                full_selection = xw.apps.active.selection.address
                current_selection = full_selection.split(':')[0].replace('$', '')
                
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
                            region_var.set(f"Region: {', '.join(range_descriptions)}")
                        else:
                            # Show count if more than 3 ranges
                            region_var.set(f"Region: {len(ranges)} Non-contiguous Ranges ({ranges[0]}, {ranges[1]}, ...)")
                            
                    elif ':' in clean_selection:
                        # Single contiguous range
                        start_cell, end_cell = clean_selection.split(':')
                        region_var.set(f"Region: Range {start_cell}:{end_cell}")
                    else:
                        # Single cell selected
                        region_var.set(f"Region: Single Cell {clean_selection}")
                except Exception as e:
                    region_var.set("Region: Single Cell")
                
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
                print(f"Error updating cell context display: {e}")
                cell_values_var.set("Above: — | Left: —")
                region_var.set("Region: Single Cell")
        
        # Configure the refresh button now that the function is defined
        refresh_context_button.config(command=update_cell_context_display)

        # Main content frame (moved to left frame)
        content_frame = ttk.Frame(left_frame)
        content_frame.pack(fill="both", expand=True)

        # Create frame for the specific data source (or all if data_source is None for backward compatibility)
        data_sources_to_process = [data_source] if data_source else self.get_all_data_sources()
        
        for current_data_source in data_sources_to_process:
            config = self.data_sources[current_data_source]
            
            # Create frame for this data source
            data_frame = ttk.LabelFrame(content_frame, text=f"Coordinates")
            data_frame.pack(side="left", fill="both", expand=True, padx=5)

            # Controls frame
            controls = ttk.Frame(data_frame)
            controls.pack(fill="x", padx=5, pady=5)

            # Label
            label = ttk.Label(controls, text=f"Select Value:")
            label.pack(side="left")

            # Combo box
            combo = ttk.Combobox(controls, values=combo_values.get(current_data_source, []), state="readonly")
            combo.pack(side="left", fill="x", expand=True, padx=5)
            
            # Store combo reference
            if current_data_source not in self.data_source_widgets:
                self.data_source_widgets[current_data_source] = {}
            self.data_source_widgets[current_data_source]['combo'] = combo

            # Button frame
            btn_frame = ttk.Frame(data_frame)
            btn_frame.pack(fill="x", padx=5)

            # Listbox (enable multi-select and preserve selection on focus change)
            listbox = tk.Listbox(data_frame, height=8, selectmode=tk.EXTENDED, exportselection=False)
            listbox.pack(fill="both", expand=True, padx=5, pady=5)
            
            # Store listbox reference
            if current_data_source not in self.data_source_widgets:
                self.data_source_widgets[current_data_source] = {}
            self.data_source_widgets[current_data_source]['listbox'] = listbox

            # Buttons - use current destination dynamically instead of capturing it
            def add_coord_wrapper():
                current_dest = self.get_current_destination()
                if not current_dest:
                    current_dest = destination  # Fallback to original destination
                add_coordinate(current_data_source, coord_entry, combo, listbox, current_dest)
            
            def remove_coord_wrapper():
                current_dest = self.get_current_destination()
                if not current_dest:
                    current_dest = destination  # Fallback to original destination
                remove_coordinate(current_data_source, listbox, current_dest)
            
            def clear_coord_wrapper():
                current_dest = self.get_current_destination()
                if not current_dest:
                    current_dest = destination  # Fallback to original destination
                clear_coordinates(current_data_source, listbox, current_dest)
            
            ttk.Button(btn_frame, text=f"Add",
                       command=add_coord_wrapper).pack(side="left", padx=2)
            ttk.Button(btn_frame, text="Remove",
                       command=remove_coord_wrapper).pack(side="left", padx=2)
            ttk.Button(btn_frame, text="Clear All",
                       command=clear_coord_wrapper).pack(side="left", padx=2)
            
            # AutoMap button with min score entry
            def automap_coordinate(dt=current_data_source, combo_box=combo):
                # Get current destination dynamically
                current_dest = self.get_current_destination()
                if not current_dest:
                    current_dest = destination  # Fallback to original destination
                
                # Get min score from right panel
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
                    self.automap_noncontiguous(dt, full_selection, combo_box, coord_entry, listbox, min_score, current_dest)
                elif ':' in clean_selection:
                    print(f"DEBUG: Single contiguous range: {clean_selection}")
                    # Single contiguous range - iterate through cells
                    self.automap_range(dt, full_selection, combo_box, coord_entry, listbox, min_score, current_dest)
                else:
                    # Single cell - perform mapping
                    print(f"DEBUG: Single cell: {clean_selection}")
                    current_coord = coord_entry.get().strip()
                    if current_coord:
                        best_match, score, header = self.auto_map_coordinate_semantic(dt, current_coord, min_score)
                        if best_match:
                            combo_box.set(best_match)
                            add_coordinate(dt, coord_entry, combo_box, listbox, current_dest)
                            header_display = f"'{header}'" if header else "None"
                            messagebox.showinfo("AutoMap Result", 
                                f"Successfully mapped {current_coord} to '{best_match}'\nSimilarity Score: {score:.3f}\nHeader: {header_display}")
                        else:
                            header_display = f"'{header}'" if header else "None"
                            messagebox.showinfo("AutoMap Result", 
                                f"No match found for {current_coord}\nBest Score: {score:.3f} (below threshold {min_score})\nHeader: {header_display}")
            
            # AutoMap button
            ttk.Button(btn_frame, text="AutoMap",
                       command=lambda dt=current_data_source, cb=combo: automap_coordinate(dt, cb)).pack(side="left", padx=2)

            # References are already stored in data_source_widgets above

        
        # Semantic mapping controls (moved to right frame)
        # Status indicator
        status_text = "Semantic model: Not loaded"
        status_color = "red"
        if self.model_loaded and self.semantic_model is not None:
            status_text = "Semantic model: Ready"
            status_color = "green"
        elif self.loading_model:
            status_text = "Semantic model: Loading..."
            status_color = "orange"
        
        self.semantic_status_label = ttk.Label(right_frame, text=status_text, foreground=status_color)
        self.semantic_status_label.pack(anchor="w", pady=(0, 5))
        print(f"DEBUG: Created semantic_status_label with text: {status_text}")
        
        # Load model button
        def load_semantic_model():
            print("DEBUG: Load Semantic Model button clicked!")
            print(f"DEBUG: model_loaded = {self.model_loaded}")
            print(f"DEBUG: loading_model = {self.loading_model}")
            
            if not self.model_loaded and not self.loading_model:
                print("DEBUG: Conditions met, starting to load model...")
                self.load_semantic_model_async()
                self.semantic_status_label.config(text="Semantic model: Loading...", foreground="orange")
            else:
                print(f"DEBUG: Conditions NOT met! model_loaded={self.model_loaded}, loading_model={self.loading_model}")
                if self.model_loaded:
                    print("DEBUG: Model already loaded!")
                    messagebox.showinfo("Info", "Semantic model is already loaded")
                elif self.loading_model:
                    print("DEBUG: Model is currently loading!")
                    messagebox.showinfo("Info", "Semantic model is currently loading, please wait...")
        
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

        # Source Sheet Configuration section (moved under Semantic Mapping)
        source_sheet_frame = ttk.Frame(right_frame, padding="5")
        source_sheet_frame.pack(fill="x", pady=(0, 10))
        
        # Source sheet name row
        source_sheet_row = tk.Frame(source_sheet_frame)
        source_sheet_row.pack(fill="x", pady=5)
        
        # Help button
        source_help_btn = tk.Button(source_sheet_row, text="?", width=2,
                                    command=lambda: self.show_help("source_sheet_name.txt"))
        source_help_btn.pack(side=tk.LEFT, padx=(5, 2))
        
        # Label
        source_sheet_label = tk.Label(source_sheet_row, text="Source Sheet Name:", width=20)
        source_sheet_label.pack(side=tk.LEFT, padx=(0, 5))
        
        # Combobox for source sheet name (allows typing and dropdown selection)
        source_sheet_entry = ttk.Combobox(source_sheet_row, state="normal")
        source_sheet_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
        # Store reference to the source sheet entry for this data source
        if not hasattr(self, 'data_source_ui_entries'):
            self.data_source_ui_entries = {}
        self.data_source_ui_entries[data_source] = {
            'source_sheet_entry': source_sheet_entry,
            'top_tag_entry': None  # Will be set when top tag entry is created
        }
        
        # Get existing sheet names and populate the combobox
        try:
            sheet_names = self.get_sheet_names()
            source_sheet_entry['values'] = sheet_names
        except Exception as e:
            print(f"Warning: Could not load sheet names for source sheet dropdown: {e}")
            source_sheet_entry['values'] = []
        
        # Set initial value
        current_source_sheet = self.data_sources[data_source]['source_sheet_name']
        # Ensure current_source_sheet is a string
        source_sheet_value = str(current_source_sheet) if current_source_sheet is not None else ""
        source_sheet_entry.set(source_sheet_value)
        
        # Bind change event to update the data source configuration
        def update_source_sheet_name(event=None):
            new_source_sheet = source_sheet_entry.get().strip()
            self.data_sources[data_source]['source_sheet_name'] = new_source_sheet
        
        # Bind events for Combobox (supports both dropdown selection and typing)
        source_sheet_entry.bind('<<ComboboxSelected>>', update_source_sheet_name)
        source_sheet_entry.bind('<KeyRelease>', update_source_sheet_name)
        source_sheet_entry.bind('<FocusOut>', update_source_sheet_name)

        # Save/Load Coordinate Maps buttons
        save_load_frame = ttk.Frame(right_frame)
        save_load_frame.pack(fill="x", pady=(0, 10))
        
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
                    for data_source in self.get_all_data_sources():
                        coordinate_data[data_source] = self.data_sources[data_source]['coordinate_values']
                    
                    # Also save coordinate conversions and combinations per data source
                    coordinate_data['conversions'] = {}
                    for data_source in self.get_all_data_sources():
                        coordinate_data['conversions'][data_source] = self.data_sources[data_source]['coordinate_conversions']
                    coordinate_data['combinations'] = {}
                    for data_source in self.get_all_data_sources():
                        coordinate_data['combinations'][data_source] = self.data_sources[data_source]['coordinate_combinations']
                    
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
                    
                    # Load coordinate values for each data source
                    for data_source in self.get_all_data_sources():
                        if data_source in coordinate_data:
                            self.data_sources[data_source]['coordinate_values'] = coordinate_data[data_source]
                    
                    # Load coordinate conversions and combinations
                    if 'conversions' in coordinate_data:
                        for data_source in self.get_all_data_sources():
                            if data_source in coordinate_data['conversions']:
                                self.data_sources[data_source]['coordinate_conversions'] = coordinate_data['conversions'][data_source]
                    if 'combinations' in coordinate_data:
                        for data_source in self.get_all_data_sources():
                            if data_source in coordinate_data['combinations']:
                                self.data_sources[data_source]['coordinate_combinations'] = coordinate_data['combinations'][data_source]
                    
                    # Update the UI
                    self.refresh_tab_content()
                    messagebox.showinfo("Success", f"Coordinate maps loaded from {file_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load coordinate maps: {e}")
        
        ttk.Button(save_load_frame, text="Save Coordinate Maps", 
                   command=save_coordinate_maps).pack(side="left", padx=5)
        ttk.Button(save_load_frame, text="Load Coordinate Maps", 
                   command=load_coordinate_maps).pack(side="left", padx=5)

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
            # Remove selected coordinates (multi-delete)
            self.coord_context_menu.add_command(label="Remove Selected", command=lambda: remove_coordinate(self.current_data_source_for_context, self.current_listbox_for_context))
            self.coord_context_menu.add_separator()
            self.coord_context_menu.add_command(label="Cancel")

        if not hasattr(self, 'selected_coord_for_context'):
            self.selected_coord_for_context = None
        if not hasattr(self, 'current_data_source_for_context'):
            self.current_data_source_for_context = None
        if not hasattr(self, 'current_listbox_for_context'):
            self.current_listbox_for_context = None

        def show_coord_context_menu(event, listbox_widget, data_source):
            # Select behavior: keep multi-selection if right-clicked item is within it; otherwise select the clicked item
            clicked_index = listbox_widget.nearest(event.y)
            current_selection = set(listbox_widget.curselection())
            if clicked_index not in current_selection:
                listbox_widget.selection_clear(0, tk.END)
                listbox_widget.selection_set(clicked_index)
                listbox_widget.activate(clicked_index)

            selected_text = listbox_widget.get(clicked_index)
            # Extract coordinate (part before ':')
            self.selected_coord_for_context = selected_text.split(':')[0].strip()
            # Store the current data source and listbox for context menu actions
            self.current_data_source_for_context = data_source
            self.current_listbox_for_context = listbox_widget

            # Check if conversion exists and enable/disable "Remove Conversion"
            if self.selected_coord_for_context in self.data_sources[data_source]['coordinate_conversions']:
                self.coord_context_menu.entryconfig("Remove Conversion", state="normal")
            else:
                self.coord_context_menu.entryconfig("Remove Conversion", state="disabled")

            # Check if combination exists and enable/disable "Remove Combination"
            if self.selected_coord_for_context in self.data_sources[data_source]['coordinate_combinations']:
                self.coord_context_menu.entryconfig("Remove Combination", state="normal")
            else:
                self.coord_context_menu.entryconfig("Remove Combination", state="disabled")

            # Enable/disable multi-remove based on selection count
            try:
                if listbox_widget.curselection():
                    self.coord_context_menu.entryconfig("Remove Selected", state="normal")
                else:
                    self.coord_context_menu.entryconfig("Remove Selected", state="disabled")
            except Exception:
                pass

            # Popup the menu
            try:
                self.coord_context_menu.tk_popup(event.x_root, event.y_root)
            finally:
                self.coord_context_menu.grab_release()

        # Bind right-click to this specific listbox
        listbox.bind("<Button-3>", lambda event, l=listbox, dt=data_source: show_coord_context_menu(event, l, dt))
        # ---

        update_listboxes()


        # Expose reinitialize method for external calls
        tab.reinitialize = reinitialize
        return tab

    def _create_filter_ui(self, parent_widget, data_source, show_info_label=True):
        """
        Consolidated helper method to create filter UI.
        Used by both create_filters_tab and set_tag_filters.
        
        Args:
            parent_widget: The parent widget (tab or window) to add UI to
            data_source: The data source name to configure filters for
            show_info_label: Whether to show the info label at the top
        """
        filters_entries = []
        combo_values = []
        
        def refresh_combo_values():
            """Re-gather combo values from the data source and update all comboboxes"""
            combo_values.clear()
            data = self.data_sources[data_source]['data']
            if data:
                # Get keys from the first entry
                first_entry = list(data.values())[0]
                if isinstance(first_entry, dict):
                    combo_values.extend(list(first_entry.keys()))
            
            # Update all existing name_entry comboboxes with new values
            for name_entry, filter_entry in filters_entries:
                current_name = name_entry.get()
                name_entry['values'] = combo_values
                # If current selection is no longer valid, clear it
                if current_name and current_name not in combo_values:
                    name_entry.set('')
                    filter_entry['values'] = []
                    filter_entry.set('')
                # If current selection is still valid, refresh filter values
                elif current_name:
                    new_filter_values = get_filter_values_for_header(current_name)
                    filter_entry['values'] = new_filter_values
                    # Clear if current filter value is no longer valid
                    if filter_entry.get() not in new_filter_values:
                        filter_entry.set('')
            
            print(f"Refreshed combo values for {data_source}: {combo_values}")
        
        def get_filter_values_for_header(header):
            """Get unique values for a specific header from the corresponding data source"""
            unique_values = set()
            data = self.data_sources[data_source]['data']
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
            self.data_sources[data_source]['tag_filters'].clear()
            for name_entry, filter_entry in filters_entries:
                name = name_entry.get()
                filter_value = filter_entry.get()
                if name:  # Allow blank filter_value for filtering empty fields
                    self.data_sources[data_source]['tag_filters'].append([name, filter_value])
            print(f"Saved Tag Filters for {data_source}: {self.data_sources[data_source]['tag_filters']}")

        # Add explanatory label if requested
        if show_info_label:
            filter_info_label = ttk.Label(parent_widget, text="Can use comma to include multiple filter terms")
            filter_info_label.pack(anchor="w", padx=5, pady=(2, 0))

        content_frame = ttk.Frame(parent_widget)
        content_frame.pack(fill="both", expand=True, padx=5, pady=5)

        button_frame = ttk.Frame(parent_widget)
        button_frame.pack(fill="x", padx=5, pady=5)

        add_button = ttk.Button(button_frame, text="Add Filter", command=lambda: add_filter_row())
        add_button.pack(side="left", padx=5)

        refresh_button = ttk.Button(button_frame, text="Refresh", command=refresh_combo_values)
        refresh_button.pack(side="left", padx=5)

        save_button = ttk.Button(button_frame, text="Save", command=save_filters)
        save_button.pack(side="right", padx=5)

        # Initial load of combo values
        refresh_combo_values()

        # Populate initial rows with existing tag filters from data source
        for name, filter_value in self.data_sources[data_source]['tag_filters']:
            add_filter_row(name, filter_value)

        add_filter_row()  # Add empty row

    def create_filters_tab(self, tab, data_source):
        """Create filters tab using the consolidated filter UI helper"""
        print("DEBUG: Starting create_filters_tab...")
        self._create_filter_ui(tab, data_source, show_info_label=True)

    def create_transform_tab(self, tab, data_source):
        print("DEBUG: Starting create_transform_tab...")
        # Source data source selector
        data_source_frame = ttk.Frame(tab)
        data_source_frame.pack(fill="x", padx=5, pady=2)

        ttk.Label(data_source_frame, text="Source data source:").pack(side="left")

        # Get available data sources for combo
        data_source_values = [dt for dt in self.get_all_data_sources()]
        data_source_combo = ttk.Combobox(data_source_frame, values=data_source_values)
        data_source_combo.pack(side="left", fill="x", expand=True, padx=5)
        
        # Set default to first data source
        if data_source_values:
            data_source_combo.set(data_source_values[0])

        # Source key selector
        key_frame = ttk.Frame(tab)
        key_frame.pack(fill="x", padx=5, pady=2)

        ttk.Label(key_frame, text="Source Key:").pack(side="left")

        key_combo = ttk.Combobox(key_frame, values=[])
        key_combo.pack(side="left", fill="x", expand=True, padx=5)

        # Update key combo when data source changes
        def update_key_combo(*args):
            selected_data_source_name = data_source_combo.get()
            data_source = selected_data_source_name  # Since we're using data source names directly now
            if data_source:
                data = self.data_sources[data_source]['data']
                if data:
                    # Get keys from the first entry
                    first_entry = list(data.values())[0]
                    if isinstance(first_entry, dict):
                        key_values = list(first_entry.keys())
                        key_combo['values'] = key_values
                        if key_values:
                            key_combo.set(key_values[0])

        data_source_combo.bind('<<ComboboxSelected>>', update_key_combo)
        update_key_combo()  # Initialize

        # Transformation code entry
        code_frame = ttk.Frame(tab)
        code_frame.pack(fill="x", padx=5, pady=5)

        ttk.Label(code_frame, text="Key, x = Index Value:").pack(side="left")
        code_entry = ttk.Entry(code_frame)
        code_entry.pack(side="left", fill="x", expand=True, padx=5)
        code_entry.insert(0, self.data_sources[data_source]['transformation_code'])

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
            selected_data_source_name = data_source_combo.get()
            transform_data_source = selected_data_source_name  # Since we're using data source names directly now
            self.data_sources[data_source]['transform_data_source'] = transform_data_source
            self.data_sources[data_source]['transform_key'] = key_combo.get()
            self.data_sources[data_source]['transformation_code'] = code_entry.get()
            print(f"Transform saved - data source: {selected_data_source_name}, Key: {self.data_sources[data_source]['transform_key']}, Code: {self.data_sources[data_source]['transformation_code']}")

        ttk.Button(tab, text="Save", command=save_transform).pack(pady=5)

    def refresh_tab_content(self, force_rebuild=False):
        """Refresh all configuration tabs within each data source tab"""
        print("DEBUG: Starting refresh_tab_content...")
        data_sources = self.get_all_data_sources()
        print(f"DEBUG: Found {len(data_sources)} data sources: {data_sources}")
        
        # Load sheet names once for all tabs to avoid multiple file access
        print("DEBUG: Loading sheet names once for all tabs...")
        sheet_names = self.get_sheet_names()
        print(f"DEBUG: Loaded {len(sheet_names)} sheet names for all tabs")
        
        # Refresh all configuration tabs within each data source tab
        for data_source in data_sources:
            print(f"DEBUG: Refreshing tabs for data_source: {data_source}")
            # Get config_notebook reference from data_source_widgets
            config_notebook = self.data_source_widgets.get(data_source, {}).get('config_notebook')
            if config_notebook:
                print(f"DEBUG: Found config notebook for {data_source}")
                
                # Only rebuild if forced or if tabs don't exist
                existing_tabs = config_notebook.winfo_children()
                if force_rebuild or len(existing_tabs) == 0:
                    print(f"DEBUG: Rebuilding tabs for {data_source}")
                    # Recreate all tabs with updated data
                    print(f"DEBUG: Destroying existing tabs for {data_source}")
                    for tab in existing_tabs:
                        tab.destroy()

                    print(f"DEBUG: Creating new tabs for {data_source}")
                    coordinates_tab = ttk.Frame(config_notebook)
                    filters_tab = ttk.Frame(config_notebook)
                    transform_tab = ttk.Frame(config_notebook)

                    config_notebook.add(coordinates_tab, text='Coordinates')
                    config_notebook.add(filters_tab, text='Filters')
                    config_notebook.add(transform_tab, text='Transform')

                    print(f"DEBUG: Creating coordinates tab for {data_source}")
                    self.create_coordinates_tab(coordinates_tab, data_source, sheet_names)
                    print(f"DEBUG: Creating filters tab for {data_source}")
                    self.create_filters_tab(filters_tab, data_source)
                    print(f"DEBUG: Creating transform tab for {data_source}")
                    self.create_transform_tab(transform_tab, data_source)
                    print(f"DEBUG: Completed tabs for {data_source}")
                else:
                    print(f"DEBUG: Skipping rebuild for {data_source} - tabs already exist")
            else:
                print(f"DEBUG: No config notebook found for {data_source}")
        
        print("DEBUG: refresh_tab_content completed")

    def sort_tabs(self):
        """Sort data source tabs alphabetically"""
        try:
            # Get current tab count
            tab_count = self.data_sources_notebook.index("end")
            
            if tab_count == 0:
                tk.messagebox.showinfo("Sort Tabs", "No tabs to sort.")
                return
            
            # Get all tab names and their current indices
            tab_info = []
            for i in range(tab_count):
                tab_name = self.data_sources_notebook.tab(i, "text")
                tab_info.append((tab_name, i))
            
            # Remember the currently selected tab
            try:
                current_tab_index = self.data_sources_notebook.index("current")
                current_tab_name = self.data_sources_notebook.tab(current_tab_index, "text")
            except:
                current_tab_name = None
            
            # Sort tab names alphabetically (case-insensitive)
            sorted_tab_info = sorted(tab_info, key=lambda x: x[0].lower())
            
            # Check if already sorted
            if tab_info == sorted_tab_info:
                tk.messagebox.showinfo("Sort Tabs", "Tabs are already sorted alphabetically.")
                return
            
            # Reorder tabs by moving them to their sorted positions
            # We need to detach and reinsert tabs in the correct order
            tabs_widgets = []
            for tab_name, _ in sorted_tab_info:
                # Find the tab widget by name
                for i in range(self.data_sources_notebook.index("end")):
                    if self.data_sources_notebook.tab(i, "text") == tab_name:
                        tab_widget = self.data_sources_notebook.nametowidget(
                            self.data_sources_notebook.tabs()[i]
                        )
                        tabs_widgets.append((tab_name, tab_widget))
                        break
            
            # Remove all tabs
            for i in range(tab_count - 1, -1, -1):
                self.data_sources_notebook.forget(i)
            
            # Re-add tabs in sorted order
            for tab_name, tab_widget in tabs_widgets:
                self.data_sources_notebook.add(tab_widget, text=tab_name)
            
            # Restore the selected tab if it was remembered
            if current_tab_name:
                for i in range(self.data_sources_notebook.index("end")):
                    if self.data_sources_notebook.tab(i, "text") == current_tab_name:
                        self.data_sources_notebook.select(i)
                        break
            
            tk.messagebox.showinfo("Sort Tabs", "Tabs have been sorted alphabetically.")
            
        except Exception as e:
            tk.messagebox.showerror("Error", f"Error sorting tabs: {str(e)}")
            print(f"Error in sort_tabs: {e}")
            import traceback
            traceback.print_exc()

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
        
        # Center the window over the main window
        center_window_over_parent(config_window)

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
        """Create a popup window to set tag filters for the primary data source using the consolidated helper"""
        # Get the current primary data source
        primary_data_source = self.get_primary_data_source()
        if not primary_data_source:
            messagebox.showerror("Error", "No data source selected")
            return
            
        view_window = tk.Toplevel(self.root)
        view_window.title(f"Set Tag Filters for {primary_data_source} (Comma for OR). ReGenerate Coordinates if necessary")

        # Use the consolidated filter UI helper
        self._create_filter_ui(view_window, primary_data_source, show_info_label=True)

        # Configure row and column weights to make them expandable
        for i in range(4):
            view_window.grid_columnconfigure(i, weight=1)

    # endregion

    # region Data Generation and Processing


    def assign_value_coordinate_to_tag(self):
        """Legacy method - now calls the dynamic version"""
        self.assign_value_coordinate_to_tag_dynamic()
    
    def assign_value_coordinate_to_tag_dynamic(self):
        """Dynamic version that works with any data sources"""
        print("Generating Coordinate-Value Data")
        self.tag_cell_values = {}  # 'a1':'LINE', 'a2':'PID' ...
        # Clear tag_cell_values_by_destination to ensure only primary data source data is used
        self.tag_cell_values_by_destination = {}
        
        def process_coordinate_value(coordinate, raw_value, data_source):
            """Helper function to process a coordinate value with conversions and combinations"""
            if raw_value is None:
                return None
                
            # Apply unit conversion first if it exists
            if coordinate in self.data_sources[data_source]['coordinate_conversions']:
                conv_details = self.data_sources[data_source]['coordinate_conversions'][coordinate]
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
        
        def apply_combinations(data, data_source):
            """Apply coordinate combinations to the data dictionary"""
            for coord, combo_info in self.data_sources[data_source]['coordinate_combinations'].items():
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
        
        # Get the primary data source (explicitly configured) - truly dynamic
        primary_data_source = self.get_primary_data_source()
        primary_data = self.data_sources[primary_data_source]['data']
        
        if not primary_data:
            print(f"No data available for primary data source: {primary_data_source}")
            return
            
        for tag in primary_data:
            if tag:
                # filter out
                continue_flag = False
                for header, filter_key in self.data_sources[primary_data_source]['tag_filters']:
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

                # Process coordinates for each destination separately
                # tag_cell_values will be structured as {destination: {tag: {coord: value}}}
                primary_data_source = self.get_primary_data_source()
                data_source_data = self.data_sources[primary_data_source]['data']
                
                # Process for each destination
                for destination in self.get_all_destinations():
                    if destination not in self.tag_cell_values_by_destination:
                        self.tag_cell_values_by_destination[destination] = {}
                    
                    data = {}
                    
                    # Access coordinate_values per destination
                    if destination not in self.data_sources[primary_data_source]['coordinate_values']:
                        self.data_sources[primary_data_source]['coordinate_values'][destination] = {}
                    coordinate_values = self.data_sources[primary_data_source]['coordinate_values'][destination]
                    
                    for coordinate, value in coordinate_values.items():
                        print(f'{primary_data_source} tag {tag}, value {value}, coord {coordinate}, destination {destination}')
                        raw_value = data_source_data[tag].get(value) # Use .get() for safety
                        processed_value = process_coordinate_value(coordinate, raw_value, primary_data_source)
                        data[coordinate] = processed_value
                        if processed_value is None:
                            print(f"  Warning: Key '{value}' not found in {primary_data_source} for tag '{tag}'. Skipping coordinate '{coordinate}'.")

                    # Apply combinations to the data
                    apply_combinations(data, primary_data_source)

                    self.tag_cell_values_by_destination[destination][tag] = data
                
                # Also maintain legacy tag_cell_values for backward compatibility (use first destination)
                # This will be used if destinations aren't being processed separately
                if not hasattr(self, 'tag_cell_values'):
                    self.tag_cell_values = {}
                if self.get_all_destinations():
                    first_dest = self.get_all_destinations()[0]
                    self.tag_cell_values[tag] = self.tag_cell_values_by_destination[first_dest][tag]

        print("Coordinate Values generated per destination:", self.tag_cell_values_by_destination)
    

    def check_halt_flag(self):
        """Callback function to check if the process should be halted"""
        return self.halt_flag
    
    def reset_halt_flag(self):
        """Reset the halt flag to False"""
        self.halt_flag = False
    
    def on_destination_tab_changed(self, event=None):
        """Update button references and coordinates listbox when destination tab changes.
        Also switches to the appropriate workbook if it's already open."""
        current_destination = self.get_current_destination()
        if current_destination and current_destination in self.destination_widgets:
            widgets = self.destination_widgets[current_destination]
            self.generate_button = widgets.get('generate_button')
            self.stop_button = widgets.get('stop_button')
            self.status_label = widgets.get('status_label')
        
        # Switch to the workbook for this destination if it exists
        if current_destination:
            dest_path, _, _, _, _, _ = self.get_destination_ui_values(current_destination)
            if dest_path:
                self.destination_datasheet = dest_path
                # Try to switch to the workbook if it's already open
                # This will switch if open, or do nothing if not open yet (will open when needed)
                try:
                    normalized_path = os.path.normpath(os.path.abspath(dest_path)).lower()
                    if normalized_path in self.excel_mgr.open_workbooks:
                        # Switch to this workbook
                        workbook_info = self.excel_mgr.open_workbooks[normalized_path]
                        try:
                            _ = workbook_info['wb'].name  # Verify it's still valid
                            print(f"DEBUG: Switching to workbook for destination: {current_destination}")
                            self.excel_mgr.wb = workbook_info['wb']
                            self.excel_mgr.app = workbook_info['app']
                            self.excel_mgr.is_dirty = workbook_info['is_dirty']
                            self.excel_mgr.original_path = workbook_info['original_path']
                            self.excel_mgr.temp_path = workbook_info['temp_path']
                            print("DEBUG: Switched to destination workbook")
                        except Exception as e:
                            print(f"DEBUG: Workbook for destination {current_destination} is no longer valid: {e}")
                            # Remove invalid workbook from cache
                            del self.excel_mgr.open_workbooks[normalized_path]
                except Exception as e:
                    print(f"DEBUG: Error switching to destination workbook: {e}")
        
        # Update coordinates listbox for the primary data source to show coordinates for the new destination
        primary_data_source = self.get_primary_data_source()
        if primary_data_source:
            listbox = self.data_source_widgets.get(primary_data_source, {}).get('listbox')
            if listbox:
                try:
                    if listbox.winfo_exists():
                        self.update_single_listbox(primary_data_source, listbox, current_destination)
                        print(f"DEBUG: Updated coordinates listbox for {primary_data_source} to show destination {current_destination}")
                except tk.TclError:
                    print(f"Warning: Could not update listbox for {primary_data_source} - widget may have been destroyed")
        
        # Update tab colors for all data sources based on the new destination
        self.update_all_tab_colors()
    
    def get_current_destination_buttons(self):
        """Get the buttons and status label for the current destination"""
        current_destination = self.get_current_destination()
        if current_destination and current_destination in self.destination_widgets:
            widgets = self.destination_widgets[current_destination]
            return (
                widgets.get('generate_button'),
                widgets.get('stop_button'),
                widgets.get('status_label')
            )
        # Fallback to instance variables if available
        return (
            getattr(self, 'generate_button', None),
            getattr(self, 'stop_button', None),
            getattr(self, 'status_label', None)
        )
    
    def set_halt_flag(self):
        """Set the halt flag to True to stop the process"""
        self.halt_flag = True
        print("Halt flag set - process will stop at next opportunity")
        _, _, status_label = self.get_current_destination_buttons()
        if status_label:
            status_label.config(text="Stopping...", fg="orange")
    
    def update_status(self, message, color="black"):
        """Update the status label if it exists"""
        _, _, status_label = self.get_current_destination_buttons()
        if status_label:
            status_label.config(text=message, fg=color)
    
    def add_datasheets(self):
        print('assigning tag coordinates')
        self.update_status("Assigning tag coordinates...", "blue")
        
        # Update instance attributes from entry widgets before processing
        self.update_instance_attributes_from_entries()
        
        # VALIDATION: Check for vital entries before proceeding
        missing_fields = []
        
        # Get primary data source info
        primary_data_source = self.get_primary_data_source()
        try:
            primary_source_sheet_name, primary_top_tag, primary_partial_match = self.get_data_source_ui_values(primary_data_source)
        except Exception as e:
            print(f"Error getting data source UI values: {e}")
            primary_source_sheet_name = ""
            primary_top_tag = ""
            primary_partial_match = False
        
        # Check top tag for primary data source
        if not primary_top_tag or not primary_top_tag.strip():
            missing_fields.append(f"Top Tag for {primary_data_source}")
        
        # Get the current destination (only process the selected one)
        current_destination = self.get_current_destination()
        if not current_destination:
            missing_fields.append("No destination selected - please select a destination tab")
        
        # Validate the current destination
        destination_validation_errors = []
        if current_destination:
            dest_path, dest_coord, dest_prefix, dest_rows, dest_sig_figs, dest_tolerance = self.get_destination_ui_values(current_destination)
            
            # Check destination path
            if not dest_path:
                destination_validation_errors.append(f"Destination '{current_destination}': Missing datasheet path")
            
            # Check rows_per_sheet, sig_figs, tolerance (should have defaults, but check anyway)
            if dest_rows < 1:
                destination_validation_errors.append(f"Destination '{current_destination}': Invalid rows_per_sheet")
        
        # If any required fields are missing, show error and cancel
        if missing_fields or destination_validation_errors:
            self.update_status("Missing required fields", "red")
            error_message = "Please fill in the following required fields before proceeding:\n\n"
            if missing_fields:
                error_message += "\n".join(f"• {field}" for field in missing_fields)
            if destination_validation_errors:
                error_message += "\n".join(f"• {field}" for field in destination_validation_errors)
            error_message += "\n\nThe Add/Update operation has been cancelled."
            messagebox.showerror("Missing Required Fields", error_message)
            return  # Cancel the operation
        
        self.assign_value_coordinate_to_tag()
        print("Adding/Updating Datasheets")
        self.update_status("Adding/Updating Datasheets...", "blue")
        
        # Reset halt flag at the start
        self.reset_halt_flag()
        self.is_processing = True
        
        # Enable stop button and disable generate button for current destination
        generate_button, stop_button, _ = self.get_current_destination_buttons()
        if stop_button:
            stop_button.config(state="normal")
        if generate_button:
            generate_button.config(state="disabled")
        
        try:
            # Get the primary data source
            primary_data_source = self.get_primary_data_source()
            
            # Get the current destination (only process the selected one)
            current_destination = self.get_current_destination()
            if not current_destination:
                self.update_status("No destination selected", "red")
                messagebox.showerror("Error", "Please select a destination tab before running Add/Update.")
                return
            
            # Get destination-specific values
            dest_path, dest_coord, dest_prefix, dest_rows, dest_sig_figs, dest_tolerance = self.get_destination_ui_values(current_destination)
            
            # Set destination_datasheet before initializing Excel (so init_excel can check if it's already open)
            if dest_path:
                self.destination_datasheet = dest_path
            
            # Initialize Excel connection (will reuse if already open for this path)
            self.init_excel()
            print('Excel initialized')
            # Additional check for valid workbook reference
            if not self.excel_mgr.wb:
                raise Exception("Excel workbook not properly initialized")
            
            # STEP 1: Get data_source SPECIFIC values directly from UI entries
            primary_source_sheet_name, primary_top_tag, primary_partial_match = self.get_data_source_ui_values(primary_data_source)
            print(f'DEBUG: data source UI values for {primary_data_source}: source_sheet={primary_source_sheet_name}, top_tag={primary_top_tag}, partial_match={primary_partial_match}')
            
            # Store the source sheet name for the report
            self.last_used_source_sheet = primary_source_sheet_name
            # Store destination-specific values for the report
            self.last_used_rows_per_sheet = dest_rows
            self.last_used_datasheet_prefix = dest_prefix
            
            print(f'Processing destination: {current_destination}')
            print(f'  Destination values:')
            print(f'    path: {dest_path}')
            print(f'    datasheet_coord: {dest_coord}')
            print(f'    ds_str: {dest_prefix}')
            print(f'    rows_per_sheet: {dest_rows}')
            print(f'    sig_figs: {dest_sig_figs}')
            print(f'    rounding_tolerance: {dest_tolerance}')
            print(f'  data_source SPECIFIC values ({primary_data_source}):')
            print(f'    top_tag: {primary_top_tag}')
            print(f'    source_sheet_name: {primary_source_sheet_name}')

            # Check for potential naming conflicts
            warning_messages = []
            
            # Check if source sheet name starts with prefix
            if primary_source_sheet_name and dest_prefix and primary_source_sheet_name.startswith(dest_prefix):
                warning_messages.append(
                    f"⚠️ Source sheet '{primary_source_sheet_name}' starts with prefix '{dest_prefix}'\n"
                    f"   This may cause the source sheet to be scanned as a datasheet."
                )
            
            # Get tag_cell_values for this destination
            if hasattr(self, 'tag_cell_values_by_destination') and current_destination in self.tag_cell_values_by_destination:
                destination_tag_cell_values = self.tag_cell_values_by_destination[current_destination]
            else:
                # Fallback to legacy tag_cell_values
                destination_tag_cell_values = self.tag_cell_values
            
            # Check if rows_per_sheet = 1 and any tag matches source sheet name
            if dest_rows == 1 and primary_source_sheet_name and primary_source_sheet_name in destination_tag_cell_values:
                warning_messages.append(
                    f"⚠️ INFO: Tag name '{primary_source_sheet_name}' matches source sheet name.\n"
                    f"   With rows_per_sheet=1, this will USE your existing source sheet directly.\n"
                    f"   Your source template will be preserved and used as-is.\n"
                    f"   If you prefer separate sheets, consider:\n"
                    f"   • Rename your source sheet to something unique (e.g., '_Template', 'Source_Template')\n"
                    f"   • Remove the tag '{primary_source_sheet_name}' from your data\n"
                    f"   • Change rows_per_sheet to a value > 1"
                )
            
            # Show warnings if any exist
            if warning_messages:
                warning_text = "⚠️ WARNING - Potential Issues Detected:\n\n" + "\n\n".join(warning_messages)
                warning_text += "\n\nDo you want to proceed anyway?"
                
                print("WARNING: Potential configuration issues detected")
                for msg in warning_messages:
                    print(msg)
                
                # Show warning dialog with Yes/No
                proceed = messagebox.askyesno("Configuration Warning", warning_text, icon='warning')
                if not proceed:
                    self.update_status("Operation cancelled by user", "orange")
                    return  # Cancel the operation

            # Convert dropdown selection to the correct parameter value
            color_option = self.color_coding_var.get()
            if color_option == "None (Black)":
                color_option = None
            # 'no_update' option is passed as-is
            print('color_option', color_option)
            print('partial_match', primary_partial_match)
            
            # Process only the current destination
            self.new_sheets, self.generation_statistics = add_update_datasheets(self.excel_mgr.wb, primary_source_sheet_name,
                                            destination_tag_cell_values, dest_coord,
                                            dest_prefix, rows_per_sheet=dest_rows,
                                            key_coordinate=primary_top_tag,
                                            sig_figs=dest_sig_figs, # Pass sig_figs
                                            tolerance=dest_tolerance, # Pass tolerance
                                            halt_callback=self.check_halt_flag, # Pass halt callback
                                            cell_update_option=color_option, # Pass color coding option
                                            partial_match=primary_partial_match, # Pass partial match option
                                            disable_green_highlight=self.disable_green_highlight_var.get(), # Pass green highlighting option
                                            fill_mode=self.fill_mode_var.get(), # Pass fill mode option
                                            append_suffix_to_green=self.append_suffix_to_green_var.get(), # Pass append suffix option
                                            clear_highlighting_on_match=self.clear_highlighting_on_match_var.get()) # Pass clear highlighting option
            self.excel_mgr.mark_as_modified()
            
            if self.halt_flag:
                print("Process was halted by user")
                self.update_status("Process halted by user", "orange")
                messagebox.showinfo("Process Halted", "The datasheet addition process was halted by the user.")
            else:
                print("DONE")
                self.update_status("Process completed successfully", "green")
                
                # Check for duplicate tags and show alert if found
                if self.generation_statistics.get('duplicate_tags'):
                    from main_functions import show_duplicate_tags_dialog
                    show_duplicate_tags_dialog(self.root, self.generation_statistics['duplicate_tags'])
                
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
            
            # Disable stop button and enable generate button for current destination
            generate_button, stop_button, _ = self.get_current_destination_buttons()
            if stop_button:
                stop_button.config(state="disabled")
            if generate_button:
                generate_button.config(state="normal")

    # endregion

    # region Data Loading and Saving

    def new_project(self):
        """Create a new project by resetting all settings to default state"""
        # Ask for confirmation
        if messagebox.askyesno("New Project", "Create a new project? All unsaved changes will be lost."):
            try:
                # Clear current settings file reference
                self.current_settings_file = None
                
                # Update window title to reflect no settings file
                self.update_window_title()
                
                # Clear data sources (without confirmation dialogs)
                self.data_sources.clear()
                self.data_source_widgets.clear()
                if hasattr(self, 'data_source_ui_entries'):
                    self.data_source_ui_entries.clear()
                
                # Clear destinations (without confirmation dialogs)
                self.destinations.clear()
                self.destination_widgets.clear()
                
                # Reset Excel manager - close all open workbooks
                try:
                    self.excel_mgr.cleanup(close_all=True)
                except Exception as e:
                    print(f"Warning: Error closing workbooks in new_project: {e}")
                
                # Reset other attributes
                self.new_sheets = []
                self.halt_flag = False
                self.is_processing = False
                self.destination_datasheet = None
                
                # Clear entries
                self.entries = []
                
                # Clean up entries and refresh UI
                self.cleanup_entries()
                self.refresh_data_sources_notebook()
                # Refresh destinations notebook to remove all tabs
                if hasattr(self, 'destinations_notebook'):
                    self.refresh_destinations_notebook()
                
                messagebox.showinfo("New Project", "New project created successfully.")
                
            except Exception as e:
                messagebox.showerror("Error", f"Failed to create new project: {str(e)}")

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
            
            # Store entry values to apply after UI rebuild
            saved_entry_values = settings_data.get('entry_values', {})
            
            # Load default settings (legacy support - no longer used)
            if 'default_settings' in settings_data:
                for key, value in settings_data['default_settings'].items():
                    # Set as instance attribute if it exists
                    if hasattr(self, key):
                        setattr(self, key, value)
            
            # Load dynamic attributes
            if 'dynamic_attributes' in settings_data:
                self.set_dynamic_attributes(settings_data['dynamic_attributes'])
            
            # Load data sources configuration
            if 'data_sources' in settings_data:
                for data_source, config in settings_data['data_sources'].items():
                    # Create the data source if it doesn't exist
                    if data_source not in self.data_sources:
                        self.add_data_source(data_source, config)
                    else:
                        # Update the existing data source configuration
                        for key, value in config.items():
                            self.data_sources[data_source][key] = value
            
            # Load destinations configuration
            if 'destinations' in settings_data:
                for destination, config in settings_data['destinations'].items():
                    # Create the destination if it doesn't exist
                    if destination not in self.destinations:
                        self.add_destination(destination, config)
                    else:
                        # Update the existing destination configuration
                        for key, value in config.items():
                            self.destinations[destination][key] = value
            else:
                # If no destinations in saved file, create a default one for backward compatibility
                if not self.destinations:
                    default_dest = "Default"
                    self.add_destination(default_dest, {
                        'path': '', 'datasheet_coord': '', 'ds_str': '',
                        'rows_per_sheet': 1, 'sig_figs': 4, 'rounding_tolerance': 1e-2
                    })
            
            # Initialize coordinate maps for all data source * destination combinations
            # This ensures coordinate maps exist even if they weren't in the saved file
            self.initialize_coordinate_maps_for_all_combinations()
            
            # Refresh the UI to reflect loaded settings
            # Clear the entries list since widgets will be recreated
            self.entries = []
            # First refresh the data sources notebook to show any new data sources
            self.refresh_data_sources_notebook()
            # Refresh destinations notebook to show any new destinations
            if hasattr(self, 'destinations_notebook'):
                self.refresh_destinations_notebook()
            # Then refresh data source frames in case new data sources were loaded
            self.refresh_data_source_frames()
            # Finally refresh all tab content to show the loaded data
            self.refresh_tab_content(force_rebuild=True)
            
            # Repopulate the entries list after UI refresh
            self.repopulate_entries_list()
            
            # NOW set entry values after UI has been rebuilt and entries repopulated
            if saved_entry_values:
                self.set_entry_values(saved_entry_values)
            
            # Set data_source specific UI entries after UI is refreshed
            if 'data_sources' in settings_data:
                for data_source, config in settings_data['data_sources'].items():
                    source_sheet_name = config.get('source_sheet_name', '')
                    top_tag = config.get('top_tag', '')
                    partial_match = config.get('partial_match', False)
                    self.set_data_source_ui_values(data_source, source_sheet_name, top_tag, partial_match)
            
            # Set destination specific UI entries after UI is refreshed
            if 'destinations' in settings_data:
                for destination, config in settings_data['destinations'].items():
                    # Update destination UI entries if they exist
                    if destination in self.destination_widgets:
                        widgets = self.destination_widgets[destination]
                        if 'datasheet_entry' in widgets:
                            widgets['datasheet_entry'].delete(0, tk.END)
                            widgets['datasheet_entry'].insert(0, config.get('path', ''))
                        if 'datasheet_coord_entry' in widgets:
                            widgets['datasheet_coord_entry'].delete(0, tk.END)
                            widgets['datasheet_coord_entry'].insert(0, config.get('datasheet_coord', ''))
                        if 'ds_str_entry' in widgets:
                            widgets['ds_str_entry'].delete(0, tk.END)
                            widgets['ds_str_entry'].insert(0, config.get('ds_str', ''))
                        if 'rows_per_sheet_entry' in widgets:
                            widgets['rows_per_sheet_entry'].delete(0, tk.END)
                            widgets['rows_per_sheet_entry'].insert(0, str(config.get('rows_per_sheet', 1)))
                        if 'sig_figs_entry' in widgets:
                            widgets['sig_figs_entry'].delete(0, tk.END)
                            widgets['sig_figs_entry'].insert(0, str(config.get('sig_figs', 4)))
                        if 'rounding_tolerance_entry' in widgets:
                            widgets['rounding_tolerance_entry'].delete(0, tk.END)
                            widgets['rounding_tolerance_entry'].insert(0, str(config.get('rounding_tolerance', 1e-2)))
            
            # Update the current settings file and window title
            self.current_settings_file = file_path
            self.save_last_settings_file_path(file_path)
            self.update_window_title()
            
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

    def migrate_to_multi_destination(self):
        """Migrate from old single-destination format to new multi-destination format"""
        try:
            # Check if migration is needed
            migration_needed = False
            migration_summary = []
            
            # Check if any data source has coordinate_values in old format (flat dict, not nested by destination)
            for data_source in self.get_all_data_sources():
                coord_vals = self.data_sources[data_source].get('coordinate_values', {})
                
                # Check if coordinate_values is a flat dict (old format)
                # New format would have coordinate_values[destination] = {...}
                if coord_vals and isinstance(coord_vals, dict):
                    # Check if it's old format: all keys are coordinate strings (like 'A1', 'B2') not destination names
                    # This is a heuristic - if we have destinations and coord_vals doesn't match that structure, it's old format
                    has_destinations = len(self.get_all_destinations()) > 0
                    if has_destinations:
                        # Check if any key in coord_vals matches a destination name
                        destinations = self.get_all_destinations()
                        is_nested = any(dest in coord_vals for dest in destinations)
                        if not is_nested and coord_vals:
                            # Likely old format - check if keys look like coordinates (A1, B2, etc.)
                            sample_key = list(coord_vals.keys())[0] if coord_vals else None
                            if sample_key and (len(sample_key) <= 4 and sample_key[0].isalpha()):
                                migration_needed = True
                                migration_summary.append(f"Data source '{data_source}': coordinate_values in old format")
                    else:
                        # No destinations exist, but coordinate_values exist - likely old format
                        if coord_vals:
                            sample_key = list(coord_vals.keys())[0] if coord_vals else None
                            if sample_key and (len(sample_key) <= 4 and sample_key[0].isalpha()):
                                migration_needed = True
                                migration_summary.append(f"Data source '{data_source}': coordinate_values in old format")
            
            # Check if global destination settings exist (old format)
            global_settings = {}
            if hasattr(self, 'datasheet_coord_entry') and self.datasheet_coord_entry:
                try:
                    coord_val = self.datasheet_coord_entry.get().strip()
                    if coord_val:
                        global_settings['datasheet_coord'] = coord_val
                        migration_needed = True
                        migration_summary.append("Global datasheet_coord found")
                except:
                    pass
            
            if hasattr(self, 'ds_str_entry') and self.ds_str_entry:
                try:
                    prefix_val = self.ds_str_entry.get().strip()
                    if prefix_val:
                        global_settings['ds_str'] = prefix_val
                        migration_needed = True
                        migration_summary.append("Global datasheet prefix found")
                except:
                    pass
            
            if hasattr(self, 'rows_per_sheet_entry') and self.rows_per_sheet_entry:
                try:
                    rows_val = self.rows_per_sheet_entry.get().strip()
                    if rows_val:
                        global_settings['rows_per_sheet'] = int(rows_val) if rows_val.isdigit() else 1
                        migration_needed = True
                        migration_summary.append("Global rows_per_sheet found")
                except:
                    pass
            
            if hasattr(self, 'sig_figs_entry') and self.sig_figs_entry:
                try:
                    sig_figs_val = self.sig_figs_entry.get().strip()
                    if sig_figs_val:
                        global_settings['sig_figs'] = int(sig_figs_val) if sig_figs_val.isdigit() else 4
                        migration_needed = True
                        migration_summary.append("Global significant figures found")
                except:
                    pass
            
            if hasattr(self, 'rounding_tolerance_entry') and self.rounding_tolerance_entry:
                try:
                    tolerance_val = self.rounding_tolerance_entry.get().strip()
                    if tolerance_val:
                        global_settings['rounding_tolerance'] = float(tolerance_val) if tolerance_val.replace('.', '').replace('-', '').isdigit() else 1e-2
                        migration_needed = True
                        migration_summary.append("Global rounding tolerance found")
                except:
                    pass
            
            if hasattr(self, 'datasheet_entry') and self.datasheet_entry:
                try:
                    path_val = self.datasheet_entry.get().strip()
                    if path_val:
                        global_settings['path'] = path_val
                        migration_needed = True
                        migration_summary.append("Global destination path found")
                except:
                    pass
            
            if not migration_needed:
                messagebox.showinfo("Migration", "No migration needed. Your data is already in the new multi-destination format.")
                return
            
            # Show confirmation dialog
            summary_text = "The following items will be migrated:\n\n" + "\n".join(f"  • {item}" for item in migration_summary)
            summary_text += "\n\nThis will:\n"
            summary_text += "  1. Create a 'Default' destination if none exists\n"
            summary_text += "  2. Migrate coordinate_values to per-destination format\n"
            summary_text += "  3. Migrate global destination settings to the Default destination\n"
            summary_text += "\nDo you want to proceed?"
            
            proceed = messagebox.askyesno("Migration to Multi-Destination Format", summary_text, icon='question')
            if not proceed:
                return
            
            # Perform migration
            migrated_items = []
            
            # Step 1: Ensure we have at least one destination
            default_dest = "Default"
            if default_dest not in self.destinations:
                self.add_destination(default_dest, {
                    'path': '',
                    'datasheet_coord': '',
                    'ds_str': '',
                    'rows_per_sheet': 1,
                    'sig_figs': 4,
                    'rounding_tolerance': 1e-2
                })
                migrated_items.append(f"Created '{default_dest}' destination")
            
            # Step 2: Migrate global settings to default destination
            if global_settings:
                for key, value in global_settings.items():
                    if key in self.destinations[default_dest]:
                        old_val = self.destinations[default_dest][key]
                        self.destinations[default_dest][key] = value
                        migrated_items.append(f"Migrated global {key}: '{old_val}' -> '{value}'")
            
            # Step 3: Migrate coordinate_values for each data source
            for data_source in self.get_all_data_sources():
                coord_vals = self.data_sources[data_source].get('coordinate_values', {})
                
                if coord_vals and isinstance(coord_vals, dict):
                    # Check if it's old format (flat dict with coordinate keys)
                    destinations = self.get_all_destinations()
                    is_nested = any(dest in coord_vals for dest in destinations)
                    
                    if not is_nested and coord_vals:
                        # Old format - migrate to new format
                        old_coord_count = len(coord_vals)
                        # Convert to nested format
                        self.data_sources[data_source]['coordinate_values'] = {
                            default_dest: coord_vals.copy()
                        }
                        migrated_items.append(f"Data source '{data_source}': Migrated {old_coord_count} coordinate mappings to '{default_dest}' destination")
            
            # Step 4: Refresh UI to show changes
            if hasattr(self, 'destinations_notebook'):
                self.refresh_destinations_notebook()
            
            # Update destination tab entries if they exist
            if default_dest in self.destination_widgets:
                widgets = self.destination_widgets[default_dest]
                if 'datasheet_entry' in widgets and global_settings.get('path'):
                    widgets['datasheet_entry'].delete(0, tk.END)
                    widgets['datasheet_entry'].insert(0, global_settings['path'])
                if 'datasheet_coord_entry' in widgets and global_settings.get('datasheet_coord'):
                    widgets['datasheet_coord_entry'].delete(0, tk.END)
                    widgets['datasheet_coord_entry'].insert(0, global_settings['datasheet_coord'])
                if 'ds_str_entry' in widgets and global_settings.get('ds_str'):
                    widgets['ds_str_entry'].delete(0, tk.END)
                    widgets['ds_str_entry'].insert(0, global_settings['ds_str'])
                if 'rows_per_sheet_entry' in widgets and global_settings.get('rows_per_sheet'):
                    widgets['rows_per_sheet_entry'].delete(0, tk.END)
                    widgets['rows_per_sheet_entry'].insert(0, str(global_settings['rows_per_sheet']))
                if 'sig_figs_entry' in widgets and global_settings.get('sig_figs'):
                    widgets['sig_figs_entry'].delete(0, tk.END)
                    widgets['sig_figs_entry'].insert(0, str(global_settings['sig_figs']))
                if 'rounding_tolerance_entry' in widgets and global_settings.get('rounding_tolerance'):
                    widgets['rounding_tolerance_entry'].delete(0, tk.END)
                    widgets['rounding_tolerance_entry'].insert(0, str(global_settings['rounding_tolerance']))
            
            # Show success message
            success_text = "Migration completed successfully!\n\nMigrated items:\n" + "\n".join(f"  • {item}" for item in migrated_items)
            success_text += "\n\nYour data is now in the new multi-destination format."
            success_text += "\nYou can now create additional destinations and configure coordinate mappings for each."
            messagebox.showinfo("Migration Complete", success_text)
            
        except Exception as e:
            error_msg = f"Error during migration: {str(e)}"
            print(error_msg)
            import traceback
            traceback.print_exc()
            messagebox.showerror("Migration Error", error_msg)

    def save_settings(self, use_pickle=True):
        """Save current settings to the current file or prompt if none exists"""
        try:
            # If no current settings file, use Save As
            if not self.current_settings_file:
                return self.save_settings_as()
            
            file_path = self.current_settings_file
            
            # Collect all settings data
            settings_data = {
                'version': '1.0',
                'saved_date': datetime.now().isoformat(),
                'entry_values': self.get_entry_values(),
                'dynamic_attributes': self.get_dynamic_attributes(),
                'data_sources': {}
            }
            
            # Save data sources configuration
            for data_source, config in self.data_sources.items():
                # Get current values from UI entries for this data source
                ui_source_sheet, ui_top_tag, ui_partial_match = self.get_data_source_ui_values(data_source)
                
                # Update config with current UI values
                config['top_tag'] = ui_top_tag
                config['source_sheet_name'] = ui_source_sheet
                config['partial_match'] = ui_partial_match
                
                settings_data['data_sources'][data_source] = {
                    'headers': config.get('headers', []),
                    'coordinate_values': config.get('coordinate_values', {}),
                    'selected_sheets': config.get('selected_sheets', None),
                    'top_tag': config.get('top_tag', ''),
                    'source_sheet_name': config.get('source_sheet_name', ''),
                    'data': config.get('data', {}),
                    'path': config.get('path', ''),
                    'partial_match': config.get('partial_match', False),
                    'tag_filters': config.get('tag_filters', []),
                    'transform_data_source': config.get('transform_data_source', None),
                    'transform_key': config.get('transform_key', None),
                    'transformation_code': config.get('transformation_code', ''),
                    'coordinate_conversions': config.get('coordinate_conversions', {}),
                    'coordinate_combinations': config.get('coordinate_combinations', {}),
                    'extraction_settings': config.get('extraction_settings', {
                        'init_tag_coord': '',
                        'init_coords_to_fields': {},
                        'tags_per_sheet': 1,
                        'selected_sheets': []
                    })
                }
            
            # Save destinations configuration - iterate through all tabs to ensure we capture all destinations
            settings_data['destinations'] = {}
            
            # First, collect all destinations from tabs in the notebook
            destinations_from_tabs = set()
            if hasattr(self, 'destinations_notebook'):
                for i in range(self.destinations_notebook.index("end")):
                    tab_text = self.destinations_notebook.tab(i, "text")
                    destinations_from_tabs.add(tab_text)
            
            # Combine destinations from tabs and dictionary to ensure we get all
            all_destinations = destinations_from_tabs.union(set(self.destinations.keys()))
            
            for destination in all_destinations:
                # Get config from dictionary if it exists, otherwise create default
                config = self.destinations.get(destination, {
                    'path': '',
                    'datasheet_coord': '',
                    'ds_str': '',
                    'rows_per_sheet': 1,
                    'sig_figs': 4,
                    'rounding_tolerance': 1e-2
                })
                
                # Get current values from UI entries for this destination
                try:
                    dest_path, dest_coord, dest_prefix, dest_rows, dest_sig_figs, dest_tolerance = self.get_destination_ui_values(destination)
                    # Update config with current UI values
                    config['path'] = dest_path
                    config['datasheet_coord'] = dest_coord
                    config['ds_str'] = dest_prefix
                    config['rows_per_sheet'] = dest_rows
                    config['sig_figs'] = dest_sig_figs
                    config['rounding_tolerance'] = dest_tolerance
                except Exception as e:
                    print(f"Warning: Could not get UI values for destination {destination}: {e}")
                    # Use stored config values as fallback
                
                settings_data['destinations'][destination] = {
                    'path': config.get('path', ''),
                    'datasheet_coord': config.get('datasheet_coord', ''),
                    'ds_str': config.get('ds_str', ''),
                    'rows_per_sheet': config.get('rows_per_sheet', 1),
                    'sig_figs': config.get('sig_figs', 4),
                    'rounding_tolerance': config.get('rounding_tolerance', 1e-2)
                }
            
            # Write to file
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(settings_data, f, indent=2, ensure_ascii=False)
            
            # Update the last settings file path and window title
            self.save_last_settings_file_path(file_path)
            self.update_window_title()
            
            messagebox.showinfo("Success", f"Settings saved successfully to:\n{file_path}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save settings: {str(e)}")
            print(f"Detailed error: {e}")
            import traceback
            traceback.print_exc()

    def save_settings_as(self, use_pickle=True):
        """Save current settings to a new JSON file (Save As)"""
        try:
            # Ask user where to save the settings file
            initial_file = f"datasheet_helper_settings_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
            if self.current_settings_file:
                # Default to the current file's directory and name
                initial_file = os.path.basename(self.current_settings_file)
                initialdir = os.path.dirname(self.current_settings_file)
            else:
                initialdir = "."
            
            file_path = filedialog.asksaveasfilename(
                title="Save Settings As",
                defaultextension=".json",
                filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
                initialfile=initial_file,
                initialdir=initialdir
            )
            
            if not file_path:
                return
            
            # Update current settings file
            self.current_settings_file = file_path
            
            # Collect all settings data
            settings_data = {
                'version': '1.0',
                'saved_date': datetime.now().isoformat(),
                'entry_values': self.get_entry_values(),
                'dynamic_attributes': self.get_dynamic_attributes(),
                'data_sources': {}
            }
            
            # Save data sources configuration
            for data_source, config in self.data_sources.items():
                # Get current values from UI entries for this data source
                ui_source_sheet, ui_top_tag, ui_partial_match = self.get_data_source_ui_values(data_source)
                
                # Update config with current UI values
                config['top_tag'] = ui_top_tag
                config['source_sheet_name'] = ui_source_sheet
                config['partial_match'] = ui_partial_match
                
                settings_data['data_sources'][data_source] = {
                    'headers': config.get('headers', []),
                    'coordinate_values': config.get('coordinate_values', {}),
                    'selected_sheets': config.get('selected_sheets', None),
                    'top_tag': config.get('top_tag', ''),
                    'source_sheet_name': config.get('source_sheet_name', ''),
                    'data': config.get('data', {}),
                    'path': config.get('path', ''),
                    'partial_match': config.get('partial_match', False),
                    'tag_filters': config.get('tag_filters', []),
                    'transform_data_source': config.get('transform_data_source', None),
                    'transform_key': config.get('transform_key', None),
                    'transformation_code': config.get('transformation_code', ''),
                    'coordinate_conversions': config.get('coordinate_conversions', {}),
                    'coordinate_combinations': config.get('coordinate_combinations', {}),
                    'extraction_settings': config.get('extraction_settings', {
                        'init_tag_coord': '',
                        'init_coords_to_fields': {},
                        'tags_per_sheet': 1,
                        'selected_sheets': []
                    })
                }
            
            # Save destinations configuration - iterate through all tabs to ensure we capture all destinations
            settings_data['destinations'] = {}
            
            # First, collect all destinations from tabs in the notebook
            destinations_from_tabs = set()
            if hasattr(self, 'destinations_notebook'):
                for i in range(self.destinations_notebook.index("end")):
                    tab_text = self.destinations_notebook.tab(i, "text")
                    destinations_from_tabs.add(tab_text)
            
            # Combine destinations from tabs and dictionary to ensure we get all
            all_destinations = destinations_from_tabs.union(set(self.destinations.keys()))
            
            for destination in all_destinations:
                # Get config from dictionary if it exists, otherwise create default
                config = self.destinations.get(destination, {
                    'path': '',
                    'datasheet_coord': '',
                    'ds_str': '',
                    'rows_per_sheet': 1,
                    'sig_figs': 4,
                    'rounding_tolerance': 1e-2
                })
                
                # Get current values from UI entries for this destination
                try:
                    dest_path, dest_coord, dest_prefix, dest_rows, dest_sig_figs, dest_tolerance = self.get_destination_ui_values(destination)
                    # Update config with current UI values
                    config['path'] = dest_path
                    config['datasheet_coord'] = dest_coord
                    config['ds_str'] = dest_prefix
                    config['rows_per_sheet'] = dest_rows
                    config['sig_figs'] = dest_sig_figs
                    config['rounding_tolerance'] = dest_tolerance
                except Exception as e:
                    print(f"Warning: Could not get UI values for destination {destination}: {e}")
                    # Use stored config values as fallback
                
                settings_data['destinations'][destination] = {
                    'path': config.get('path', ''),
                    'datasheet_coord': config.get('datasheet_coord', ''),
                    'ds_str': config.get('ds_str', ''),
                    'rows_per_sheet': config.get('rows_per_sheet', 1),
                    'sig_figs': config.get('sig_figs', 4),
                    'rounding_tolerance': config.get('rounding_tolerance', 1e-2)
                }
            
            # Write to file
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(settings_data, f, indent=2, ensure_ascii=False)
            
            # Update the last settings file path and window title
            self.save_last_settings_file_path(file_path)
            self.update_window_title()
            
            messagebox.showinfo("Success", f"Settings saved successfully to:\n{file_path}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save settings: {str(e)}")
            print(f"Detailed error: {e}")
            import traceback
            traceback.print_exc()

    def update_window_title(self):
        """Update the window title to show the current settings file path"""
        base_title = "Datasheet Helper App"
        if self.current_settings_file:
            # Show just the filename and directory for brevity
            filename = os.path.basename(self.current_settings_file)
            self.root.title(f"{base_title} - [{filename}]")
        else:
            self.root.title(base_title)
    
    def save_last_settings_file_path(self, file_path):
        """Save the path to the last settings file for auto-loading on next startup"""
        try:
            with open(self.last_settings_file_path, 'w', encoding='utf-8') as f:
                f.write(file_path)
        except Exception as e:
            print(f"Error saving last settings file path: {e}")
    
    def load_last_settings_file_path(self):
        """Load the path to the last settings file"""
        try:
            if os.path.exists(self.last_settings_file_path):
                with open(self.last_settings_file_path, 'r', encoding='utf-8') as f:
                    return f.read().strip()
        except Exception as e:
            print(f"Error loading last settings file path: {e}")
        return None
    
    def auto_load_last_settings(self):
        """Auto-load the last settings file if it exists"""
        last_file = self.load_last_settings_file_path()
        if last_file and os.path.exists(last_file):
            try:
                print(f"DEBUG: Auto-loading last settings file: {last_file}")
                with open(last_file, 'r', encoding='utf-8') as f:
                    settings_data = json.load(f)
                
                # Validate the settings file format
                if 'version' not in settings_data:
                    print("DEBUG: Invalid settings file format, skipping auto-load")
                    return
                
                # Store entry values to apply after UI rebuild
                saved_entry_values = settings_data.get('entry_values', {})
                
                # Load default settings (legacy support - no longer used)
                if 'default_settings' in settings_data:
                    for key, value in settings_data['default_settings'].items():
                        # Set as instance attribute if it exists
                        if hasattr(self, key):
                            setattr(self, key, value)
                
                # Load dynamic attributes
                if 'dynamic_attributes' in settings_data:
                    self.set_dynamic_attributes(settings_data['dynamic_attributes'])
                
                # Load data sources configuration
                if 'data_sources' in settings_data:
                    for data_source, config in settings_data['data_sources'].items():
                        # Create the data source if it doesn't exist
                        if data_source not in self.data_sources:
                            self.add_data_source(data_source, config)
                        else:
                            # Update the existing data source configuration
                            for key, value in config.items():
                                self.data_sources[data_source][key] = value
                
                # Load destinations configuration
                if 'destinations' in settings_data:
                    for destination, config in settings_data['destinations'].items():
                        # Create the destination if it doesn't exist
                        if destination not in self.destinations:
                            self.add_destination(destination, config)
                        else:
                            # Update the existing destination configuration
                            for key, value in config.items():
                                self.destinations[destination][key] = value
                else:
                    # If no destinations in saved file, create a default one for backward compatibility
                    if not self.destinations:
                        default_dest = "Default"
                        self.add_destination(default_dest, {
                            'path': '', 'datasheet_coord': '', 'ds_str': '',
                            'rows_per_sheet': 1, 'sig_figs': 4, 'rounding_tolerance': 1e-2
                        })
                
                # Initialize coordinate maps for all data source * destination combinations
                # This ensures coordinate maps exist even if they weren't in the saved file
                self.initialize_coordinate_maps_for_all_combinations()
                
                # Refresh the UI to reflect loaded settings
                # Clear the entries list since widgets will be recreated
                self.entries = []
                # First refresh the data sources notebook to show any new data sources
                self.refresh_data_sources_notebook()
                # Refresh destinations notebook to show any new destinations
                if hasattr(self, 'destinations_notebook'):
                    self.refresh_destinations_notebook()
                # Then refresh data source frames in case new data sources were loaded
                self.refresh_data_source_frames()
                # Finally refresh all tab content to show the loaded data
                self.refresh_tab_content(force_rebuild=True)
                
                # Repopulate the entries list after UI refresh
                self.repopulate_entries_list()
                
                # NOW set entry values after UI has been rebuilt and entries repopulated
                if saved_entry_values:
                    self.set_entry_values(saved_entry_values)
                
                # Set data_source specific UI entries after UI is refreshed
                if 'data_sources' in settings_data:
                    for data_source, config in settings_data['data_sources'].items():
                        source_sheet_name = config.get('source_sheet_name', '')
                        top_tag = config.get('top_tag', '')
                        partial_match = config.get('partial_match', False)
                        self.set_data_source_ui_values(data_source, source_sheet_name, top_tag, partial_match)
                
                # Set destination specific UI entries after UI is refreshed
                if 'destinations' in settings_data:
                    for destination, config in settings_data['destinations'].items():
                        # Update destination UI entries if they exist
                        if destination in self.destination_widgets:
                            widgets = self.destination_widgets[destination]
                            if 'datasheet_entry' in widgets:
                                widgets['datasheet_entry'].delete(0, tk.END)
                                widgets['datasheet_entry'].insert(0, config.get('path', ''))
                            if 'datasheet_coord_entry' in widgets:
                                widgets['datasheet_coord_entry'].delete(0, tk.END)
                                widgets['datasheet_coord_entry'].insert(0, config.get('datasheet_coord', ''))
                            if 'ds_str_entry' in widgets:
                                widgets['ds_str_entry'].delete(0, tk.END)
                                widgets['ds_str_entry'].insert(0, config.get('ds_str', ''))
                            if 'rows_per_sheet_entry' in widgets:
                                widgets['rows_per_sheet_entry'].delete(0, tk.END)
                                widgets['rows_per_sheet_entry'].insert(0, str(config.get('rows_per_sheet', 1)))
                            if 'sig_figs_entry' in widgets:
                                widgets['sig_figs_entry'].delete(0, tk.END)
                                widgets['sig_figs_entry'].insert(0, str(config.get('sig_figs', 4)))
                            if 'rounding_tolerance_entry' in widgets:
                                widgets['rounding_tolerance_entry'].delete(0, tk.END)
                                widgets['rounding_tolerance_entry'].insert(0, str(config.get('rounding_tolerance', 1e-2)))
                
                # Update the current settings file and window title
                self.current_settings_file = last_file
                self.update_window_title()
                
                print(f"DEBUG: Auto-loaded settings from: {last_file}")
                
            except Exception as e:
                print(f"Error auto-loading settings: {e}")
                import traceback
                traceback.print_exc()
                # Don't show error message to user on auto-load failure, just continue with default state

    def get_entry_values(self):
        """Get current values from all entry widgets"""
        entry_values = {}
        
        print(f'DEBUG: Found {len(self.entries)} entries')
        for entry, variable in self.entries:
            print(f'DEBUG: Processing entry for variable: {variable}')
        
        # Get values from self.entries (main entry widgets)
        for entry, variable in self.entries:
            try:
                # Handle different widget types
                if hasattr(entry, 'get'):
                    value = entry.get()
                    entry_values[variable] = value
                    print(f'DEBUG: Got value for {variable}: "{value}"')
                elif hasattr(entry, 'selection_get'):
                    # Handle text widgets
                    entry_values[variable] = entry.selection_get()
                else:
                    entry_values[variable] = str(entry)
            except Exception as e:
                print(f"Error getting entry value for {variable}: {e}")
                entry_values[variable] = ""
        
        # Add data source paths from config
        for data_source in self.get_all_data_sources():
            path_value = self.data_sources[data_source]['path']
            entry_values[data_source] = path_value
            print(f'DEBUG: Added data source path for {data_source}: "{path_value}"')
        
        # No fallbacks - we should get all values from UI entries
        required_entries = ['datasheet_coord', 'ds_str', 'rows_per_sheet', 'sig_figs', 'rounding_tolerance']
        missing_after_search = [var for var in required_entries if var not in entry_values]
        
        if missing_after_search:
            print(f'ERROR: Still missing entries after search: {missing_after_search}')
            print('ERROR: Cannot proceed without UI entry values')
        
        return entry_values
    
    def repopulate_entries_list(self):
        """Repopulate the entries list by finding all entry widgets in the UI"""
        print("DEBUG: Repopulating entries list...")
        self.entries = []
        
        # Find all entry widgets in the main window
        self._find_entry_widgets(self.root)
        
        print(f"DEBUG: Repopulated entries list with {len(self.entries)} entries")
        for entry, variable in self.entries:
            print(f"DEBUG: Found entry for variable: {variable}")
    
    def _find_entry_widgets(self, widget):
        """Recursively find all entry widgets and add them to the entries list"""
        # Check if this widget is an entry widget with a known variable
        if hasattr(widget, 'get') and hasattr(widget, '_variable_name'):
            # This is an entry widget with a stored variable name
            self.entries.append((widget, widget._variable_name))
        elif hasattr(widget, 'get') and hasattr(widget, '_data_source'):
            # This is a data source entry widget
            self.entries.append((widget, widget._data_source))
        
        # Recursively check all children
        for child in widget.winfo_children():
            self._find_entry_widgets(child)
    
    def set_entry_values(self, entry_values):
        """Set values to all entry widgets"""
        print(f"\nDEBUG set_entry_values: Setting {len(entry_values)} values")
        print(f"DEBUG set_entry_values: Available entries: {len(self.entries)}")
        for entry, var in self.entries:
            print(f"  - Entry for: {var}")
        
        for variable, value in entry_values.items():
            print(f"\nDEBUG: Trying to set '{variable}' = '{value}'")
            try:
                found = False
                # Find the corresponding entry widget
                for entry, var in self.entries:
                    if var == variable:
                        found = True
                        print(f"  -> Found matching entry widget for '{variable}'")
                        # Handle different widget types
                        if hasattr(entry, 'delete') and hasattr(entry, 'insert'):
                            # Standard Entry widget
                            entry.delete(0, tk.END)
                            entry.insert(0, str(value))
                            print(f"  -> Set entry widget value to '{value}'")
                        elif hasattr(entry, 'set'):
                            # Combobox widget
                            entry.set(str(value))
                            print(f"  -> Set combobox value to '{value}'")
                        elif hasattr(entry, 'delete') and hasattr(entry, 'insert'):
                            # Text widget
                            entry.delete(1.0, tk.END)
                            entry.insert(1.0, str(value))
                            print(f"  -> Set text widget value to '{value}'")
                        break
                
                if not found:
                    # Check if this is a data source variable
                    if variable in self.data_sources:
                        print(f"  -> Not found in entries, but is a data source. Setting config.")
                        self.data_sources[variable]['path'] = value
                    else:
                        print(f"  -> WARNING: Variable '{variable}' not found in entries or data sources!")
            except Exception as e:
                print(f"Error setting entry value for {variable}: {e}")
                import traceback
                traceback.print_exc()
    
    def get_dynamic_attributes(self):
        """Get any additional dynamic attributes that are important for the application state"""
        dynamic_attrs = {}
        
        # Get attributes that might not be in default_settings but are important
        important_attrs = [
            'current_transform_data_source', 'current_transform_key',
            'tag_filters', 'tag_cell_values', 'coordinate_conversions', 
            'coordinate_combinations'
            # NOTE: semantic_model, model_loaded, loading_model are excluded
            # because the ML model cannot be serialized and should be reloaded each session
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
        
        # Safety check: Ensure semantic model state is consistent
        # (model_loaded and loading_model should not be restored from saved state)
        if hasattr(self, 'model_loaded') and self.model_loaded:
            if not hasattr(self, 'semantic_model') or self.semantic_model is None:
                print("DEBUG: Detected inconsistent semantic model state after restore. Resetting flags.")
                self.model_loaded = False
                self.loading_model = False
    
    def _is_serializable(self, obj):
        """Check if an object is JSON serializable"""
        try:
            json.dumps(obj)
            return True
        except (TypeError, ValueError):
            return False

    
    def load_data_source_from_datasheet(self, data_source, file_path=None):
        """Load data for any data source from datasheet"""
        config = self.data_sources[data_source]
        name = data_source
        
        def set_data(data):
            self.data_sources[data_source]['data'] = data
            # Update the tab indicator to show data status
            self.add_checkmark_to_tab(data_source)
            # Update coordinates section combo box with new data
            self.update_coordinates_combo_box(data_source)
            self.refresh_tab_content()
            print(f"{name} data loaded from datasheet")
        
        app_window = tk.Toplevel(self.root)
        DatasheetExtractor(app_window, callback=set_data, file_path=file_path, data_source=data_source, main_app=self)

    def load_data_source_from_json(self, data_source):
        """Load data for any data source from JSON file"""
        config = self.data_sources[data_source]
        name = data_source
        
        # Ask the user to select a JSON file
        file_path = filedialog.askopenfilename(
            title=f"Select JSON file for {name}",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        
        if file_path:
            data = load_dict_from_json(file_path)
            self.data_sources[data_source]['data'] = data
            # Update the tab indicator to show data status
            self.add_checkmark_to_tab(data_source)
            # Update coordinates section combo box with new data
            self.update_coordinates_combo_box(data_source)
            self.refresh_tab_content()
            print(f"{name} data loaded from JSON")

    def update_history(self, history_list, new_value, max_items=7):
        """
        Update a history list with a new value, keeping only the last max_items.
        Most recent items appear first in the list.
        Only updates if the value is not already in the history.
        
        Args:
            history_list: The list to update
            new_value: The new value to add
            max_items: Maximum number of items to keep (default 7)
        """
        # Only add if it's a new value (not already in history)
        if new_value in history_list:
            return  # Don't modify history if value already exists
        
        # Insert at the beginning (most recent first)
        history_list.insert(0, new_value)
        
        # Trim to max_items
        while len(history_list) > max_items:
            history_list.pop()
    
    def update_data_source_keys(self, data_source, parent=None, refresh_callback=None):
        """Update keys for any data source using transformation code"""
        config = self.data_sources[data_source]
        name = data_source
        
        # Default suggestion
        default_code = '"-".join(x.split("-")[-2:])'
        
        # Build options list with history first, then default if not in history
        options = self.modify_keys_history.copy()
        if default_code not in options:
            options.append(default_code)
        
        code = ask_combobox(f"Modify Keys - {name}", 
                           f"Enter transformation code for {name}:",
                           options=options,
                           parent=parent,
                           initialvalue=options[0] if options else default_code)
        
        if code:
            # Update history
            self.update_history(self.modify_keys_history, code)
            
            data = self.data_sources[data_source]['data']
            transformed_data = transform_dictionary(data, code)
            self.data_sources[data_source]['data'] = transformed_data
            
            # Update tab text to show checkmark
            self.add_checkmark_to_tab(data_source)
            
            print(f"{name} keys updated")
            
            # Refresh the display if callback provided
            if refresh_callback:
                refresh_callback()

    # endregion

    # region Data Display

    def view_data(self, data_source=None, initial_search=""):
        """View data for a data source or special case"""

        # New approach - data_source parameter
        if data_source and data_source in self.data_sources:
            print(f"Viewing {data_source}")
            self.display_data_source(data_source, initial_search=initial_search)
        else:
            print(f"Unknown data source: {data_source}")
    
    def display_data_source(self, data_source, initial_search=""):
        """Display data for any data source using the centralized system"""
        config = self.data_sources[data_source]
        name = data_source
        data = self.data_sources[data_source]['data']
        
        # Create a new window
        view_window = tk.Toplevel(self.root)
        view_window.title(name)
        view_window.transient(self.root)
        # Removed grab_set() to allow opening multiple viewers and interacting with search results
        
        dialog_width = 900  # Set a reasonable default width
        dialog_height = 600  # Set a reasonable default height
        view_window.geometry(f"{dialog_width}x{dialog_height}")
        
        # Center the dialog over the main window
        center_window_over_parent(view_window)

        # Create a search frame
        search_frame = tk.Frame(view_window)
        search_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # Search label and entry
        search_label = tk.Label(search_frame, text="Search:")
        search_label.pack(side=tk.LEFT, padx=(0, 5))
        
        search_var = tk.StringVar(value=initial_search)  # Set initial search value
        search_entry = tk.Entry(search_frame, textvariable=search_var, width=30)
        search_entry.pack(side=tk.LEFT, padx=5)
        
        # Match counter label
        match_label = tk.Label(search_frame, text="")
        match_label.pack(side=tk.LEFT, padx=10)
        
        # Navigation buttons
        prev_button = tk.Button(search_frame, text="◀ Previous", width=10)
        prev_button.pack(side=tk.LEFT, padx=2)
        
        next_button = tk.Button(search_frame, text="Next ▶", width=10)
        next_button.pack(side=tk.LEFT, padx=2)

        # Create a button frame for actions
        button_frame = tk.Frame(view_window)
        button_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # Placeholder for search reapply function (will be set later)
        reapply_search = [None]

        def refresh_display():
            """Refresh the display with current data"""
            scrolled_text.configure(state='normal')
            scrolled_text.delete(1.0, tk.END)
            
            current_data = self.data_sources[data_source]['data']
            if current_data:
                for key, value in current_data.items():
                    scrolled_text.insert(tk.END, f"{key}: {value}\n\n")
            else:
                scrolled_text.insert(tk.END, f"No {name} data available.")
            
            # Keep it editable - removed state='disabled'
            
            # Reapply search highlighting if function is available
            if reapply_search[0] is not None:
                reapply_search[0]()

        def split_keys():
            """Split composite keys based on delimiter"""
            # Common delimiters
            common_delimiters = [";", ",", "|", "-", "_", ":", "/"]
            
            # Build options list with history first, then common delimiters not in history
            options = self.split_keys_history.copy()
            for delim in common_delimiters:
                if delim not in options:
                    options.append(delim)
            
            # Ask for delimiter
            delimiter = ask_combobox("Split Keys", 
                                    "Enter delimiter to split keys on:",
                                    options=options,
                                    parent=view_window,
                                    initialvalue=options[0] if options else ";")
            
            if not delimiter:
                return
            
            # Update history
            self.update_history(self.split_keys_history, delimiter)
            
            # Get current data
            current_data = self.data_sources[data_source]['data']
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
            self.data_sources[data_source]['data'] = new_data
            
            # Update tab text to show checkmark
            self.add_checkmark_to_tab(data_source)
            
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
                    self.data_sources[data_source]['data'] = pasted_data
                    
                    # Update coordinates combo box with new data
                    self.update_coordinates_combo_box(data_source)
                    
                    # Update tab text to show checkmark
                    self.add_checkmark_to_tab(data_source)
                    
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

        def copy_to_clipboard():
            """Copy current data as JSON to clipboard"""
            try:
                current_data = self.data_sources[data_source]['data']
                if not current_data:
                    tk.messagebox.showwarning("No Data", f"No {name} data available to copy.")
                    return
                
                # Convert data to JSON string with indentation for readability
                json_string = json.dumps(current_data, indent=2, ensure_ascii=False)
                
                # Copy to clipboard
                self.root.clipboard_clear()
                self.root.clipboard_append(json_string)
                
                # Show success message
                tk.messagebox.showinfo("Copy Complete", 
                                     f"Successfully copied {len(current_data)} entries to clipboard as JSON.")
                
            except Exception as e:
                tk.messagebox.showerror("Copy Error", 
                                      f"An error occurred while copying to clipboard:\n{str(e)}")

        def copy_as_table():
            """Copy current data as a table with headers from first entry's keys"""
            try:
                current_data = self.data_sources[data_source]['data']
                if not current_data:
                    tk.messagebox.showwarning("No Data", f"No {name} data available to copy.")
                    return
                
                # Get the first entry to determine headers
                first_key = next(iter(current_data))
                first_value = current_data[first_key]
                
                # Check if first value is a dictionary
                if not isinstance(first_value, dict):
                    tk.messagebox.showwarning("Invalid Format", 
                                            "The first entry's value must be a dictionary to create a table.")
                    return
                
                # Get headers from the first entry's keys
                headers = list(first_value.keys())
                if not headers:
                    tk.messagebox.showwarning("No Headers", 
                                            "The first entry has no keys to use as headers.")
                    return
                
                # Sanitize headers (remove tabs, newlines, carriage returns)
                sanitized_headers = []
                for h in headers:
                    header_str = str(h)
                    # Replace tabs and newlines to avoid breaking table structure
                    header_str = header_str.replace('\t', ' ').replace('\n', ' ').replace('\r', ' ')
                    sanitized_headers.append(header_str)
                
                # Build the table
                table_lines = []
                
                # Add header row (tab-separated)
                table_lines.append('\t'.join(sanitized_headers))
                
                # Add data rows
                for key, value in current_data.items():
                    if isinstance(value, dict):
                        # Extract values in the same order as headers
                        row_values = []
                        for header in headers:
                            cell_value = value.get(header, '')
                            # Convert to string and handle None
                            if cell_value is None:
                                cell_value = ''
                            else:
                                cell_value = str(cell_value)
                            # Replace tabs and newlines to avoid breaking table structure
                            cell_value = cell_value.replace('\t', ' ').replace('\n', ' ').replace('\r', ' ')
                            row_values.append(cell_value)
                        table_lines.append('\t'.join(row_values))
                    else:
                        # If value is not a dict, create a single-column row
                        cell_value = str(value).replace('\t', ' ').replace('\n', ' ').replace('\r', ' ')
                        table_lines.append(cell_value)
                
                # Join all lines with newlines
                table_string = '\n'.join(table_lines)
                
                # Copy to clipboard
                self.root.clipboard_clear()
                self.root.clipboard_append(table_string)
                
                # Show success message
                tk.messagebox.showinfo("Copy Complete", 
                                     f"Successfully copied {len(current_data)} rows to clipboard as table.\n"
                                     f"Headers: {', '.join(headers)}")
                
            except Exception as e:
                tk.messagebox.showerror("Copy Error", 
                                      f"An error occurred while copying to clipboard:\n{str(e)}")

        def sort_data():
            """Sort the dictionary data by keys or values"""
            current_data = self.data_sources[data_source]['data']
            if not current_data:
                tk.messagebox.showwarning("No Data", f"No {name} data available to sort.")
                return
            
            # Create a dialog to choose sort options
            sort_dialog = tk.Toplevel(view_window)
            sort_dialog.title("Sort Data")
            sort_dialog.transient(view_window)
            sort_dialog.grab_set()
            center_window_over_parent(sort_dialog)
            # Center the dialog
            dialog_width = 400
            dialog_height = 300
            sort_dialog.geometry(f"{dialog_width}x{dialog_height}")
            
            # Sort by option
            sort_frame = tk.LabelFrame(sort_dialog, text="Sort Options", padx=10, pady=10)
            sort_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
            
            sort_by_var = tk.StringVar(value="key")
            tk.Radiobutton(sort_frame, text="Sort by Keys (A-Z)", variable=sort_by_var, value="key").pack(anchor=tk.W)
            tk.Radiobutton(sort_frame, text="Sort by Keys (Z-A)", variable=sort_by_var, value="key_reverse").pack(anchor=tk.W)
            tk.Radiobutton(sort_frame, text="Sort by Values (A-Z)", variable=sort_by_var, value="value").pack(anchor=tk.W)
            tk.Radiobutton(sort_frame, text="Sort by Values (Z-A)", variable=sort_by_var, value="value_reverse").pack(anchor=tk.W)
            
            # Add option to sort by field in value
            field_sort_frame = tk.Frame(sort_frame)
            field_sort_frame.pack(anchor=tk.W, pady=5)
            tk.Radiobutton(field_sort_frame, text="Sort by Field in Value:", variable=sort_by_var, value="field").pack(side=tk.LEFT)
            field_entry = tk.Entry(field_sort_frame, width=20)
            field_entry.pack(side=tk.LEFT, padx=5)
            
            # Add option for reverse field sort
            field_sort_reverse_frame = tk.Frame(sort_frame)
            field_sort_reverse_frame.pack(anchor=tk.W, pady=2)
            tk.Radiobutton(field_sort_reverse_frame, text="Sort by Field in Value (Z-A):", variable=sort_by_var, value="field_reverse").pack(side=tk.LEFT)
            
            # Helper text
            tk.Label(sort_frame, text="(e.g., 'name' or 'data.value' for nested)", 
                    font=('Arial', 8), fg='gray').pack(anchor=tk.W, padx=20)
            
            def get_nested_value(obj, field_path):
                """Get value from nested dictionary using dot notation"""
                keys = field_path.split('.')
                value = obj
                for key in keys:
                    if isinstance(value, dict):
                        value = value.get(key)
                        if value is None:
                            return None
                    else:
                        return None
                return value
            
            def apply_sort():
                sort_option = sort_by_var.get()
                
                try:
                    if sort_option == "key":
                        sorted_data = dict(sorted(current_data.items(), key=lambda item: str(item[0]).lower()))
                    elif sort_option == "key_reverse":
                        sorted_data = dict(sorted(current_data.items(), key=lambda item: str(item[0]).lower(), reverse=True))
                    elif sort_option == "value":
                        sorted_data = dict(sorted(current_data.items(), key=lambda item: str(item[1]).lower()))
                    elif sort_option == "value_reverse":
                        sorted_data = dict(sorted(current_data.items(), key=lambda item: str(item[1]).lower(), reverse=True))
                    elif sort_option == "field" or sort_option == "field_reverse":
                        field_name = field_entry.get().strip()
                        if not field_name:
                            tk.messagebox.showerror("Field Required", 
                                                  "Please enter a field name to sort by.")
                            return
                        
                        # Check if any values are dictionaries with the specified field
                        has_field = False
                        missing_field_count = 0
                        
                        for value in current_data.values():
                            field_value = get_nested_value(value, field_name)
                            if field_value is not None:
                                has_field = True
                            else:
                                missing_field_count += 1
                        
                        if not has_field:
                            tk.messagebox.showerror("Field Not Found", 
                                                  f"Field '{field_name}' not found in any values.")
                            return
                        
                        # Sort by the field, putting items without the field at the end
                        def sort_key(item):
                            field_value = get_nested_value(item[1], field_name)
                            if field_value is None:
                                return (1, "")  # Put None values at the end
                            return (0, str(field_value).lower())
                        
                        sorted_data = dict(sorted(current_data.items(), 
                                                key=sort_key,
                                                reverse=(sort_option == "field_reverse")))
                        
                        if missing_field_count > 0:
                            tk.messagebox.showinfo("Note", 
                                                 f"{missing_field_count} entries did not have the field '{field_name}' "
                                                 f"and were placed at the end.")
                    
                    # Update the data with sorted version
                    self.data_sources[data_source]['data'] = sorted_data
                    
                    # Update tab text to show checkmark
                    self.add_checkmark_to_tab(data_source)
                    
                    # Refresh the display
                    refresh_display()
                    
                    # Close the sort dialog
                    sort_dialog.destroy()
                    
                    tk.messagebox.showinfo("Sort Complete", 
                                         f"Successfully sorted {len(sorted_data)} entries.")
                    
                except Exception as e:
                    tk.messagebox.showerror("Sort Error", 
                                          f"An error occurred while sorting:\n{str(e)}")
            
            # Buttons
            button_frame_sort = tk.Frame(sort_dialog)
            button_frame_sort.pack(fill=tk.X, padx=10, pady=5)
            
            tk.Button(button_frame_sort, text="Apply", command=apply_sort).pack(side=tk.LEFT, padx=5)
            tk.Button(button_frame_sort, text="Cancel", command=sort_dialog.destroy).pack(side=tk.LEFT, padx=5)

        def layered_sort_data():
            """Apply multiple sort layers to the dictionary data"""
            current_data = self.data_sources[data_source]['data']
            if not current_data:
                tk.messagebox.showwarning("No Data", f"No {name} data available to sort.")
                return
            
            # Create a dialog for layered sorting
            layered_sort_dialog = tk.Toplevel(view_window)
            layered_sort_dialog.title("Layered Sort Data")
            layered_sort_dialog.transient(view_window)
            layered_sort_dialog.grab_set()
            center_window_over_parent(layered_sort_dialog)
            
            dialog_width = 600
            dialog_height = 650
            layered_sort_dialog.geometry(f"{dialog_width}x{dialog_height}")
            
            # Main frame
            main_frame = tk.Frame(layered_sort_dialog)
            main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
            
            # Instructions
            instructions = tk.Label(main_frame, 
                                   text="Create multiple sort layers. Each layer will be applied in order (1st layer has highest priority).",
                                   font=('Arial', 9), fg='blue')
            instructions.pack(anchor=tk.W, pady=(0, 10))
            
            # Sort layers frame
            layers_frame = tk.LabelFrame(main_frame, text="Sort Layers", padx=10, pady=10)
            layers_frame.pack(fill=tk.BOTH, expand=True)
            
            # Listbox to show sort layers
            layers_listbox = tk.Listbox(layers_frame, height=8)
            layers_listbox.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
            
            # Scrollbar for listbox
            layers_scrollbar = tk.Scrollbar(layers_frame, orient=tk.VERTICAL, command=layers_listbox.yview)
            layers_listbox.configure(yscrollcommand=layers_scrollbar.set)
            
            # Sort layers storage
            sort_layers = []
            
            def get_nested_value(obj, field_path):
                """Get value from nested dictionary using dot notation"""
                keys = field_path.split('.')
                value = obj
                for key in keys:
                    if isinstance(value, dict):
                        value = value.get(key)
                        if value is None:
                            return None
                    else:
                        return None
                return value
            
            def update_layers_display():
                """Update the listbox display of sort layers"""
                layers_listbox.delete(0, tk.END)
                for i, layer in enumerate(sort_layers):
                    layer_text = f"{i+1}. {layer['type']} ({'Descending' if layer['reverse'] else 'Ascending'})"
                    if layer['type'] == 'field' and layer['field']:
                        layer_text += f" - Field: {layer['field']}"
                    elif layer['type'] == 'custom':
                        layer_text += f" - Pattern: {layer['pattern']}"
                        if layer['custom_sort_by'] == 'field' and layer['custom_field']:
                            layer_text += f" (Field: {layer['custom_field']})"
                        elif layer['custom_sort_by'] == 'key':
                            layer_text += " (Keys)"
                        elif layer['custom_sort_by'] == 'value':
                            layer_text += " (Values)"
                    layers_listbox.insert(tk.END, layer_text)
            
            def add_sort_layer():
                """Add a new sort layer"""
                layer_dialog = tk.Toplevel(layered_sort_dialog)
                layer_dialog.title("Add Sort Layer")
                layer_dialog.transient(layered_sort_dialog)
                layer_dialog.grab_set()
                layer_dialog.geometry("400x400")
                center_window_over_parent(layer_dialog)
                
                # Sort type selection
                type_frame = tk.LabelFrame(layer_dialog, text="Sort Type", padx=10, pady=10)
                type_frame.pack(fill=tk.X, padx=10, pady=10)
                
                sort_type_var = tk.StringVar(value="key")
                tk.Radiobutton(type_frame, text="Sort by Keys", variable=sort_type_var, value="key").pack(anchor=tk.W)
                tk.Radiobutton(type_frame, text="Sort by Values", variable=sort_type_var, value="value").pack(anchor=tk.W)
                tk.Radiobutton(type_frame, text="Sort by Field in Value", variable=sort_type_var, value="field").pack(anchor=tk.W)
                tk.Radiobutton(type_frame, text="Custom Sort Pattern", variable=sort_type_var, value="custom").pack(anchor=tk.W)
                
                # Field entry for field sorting
                field_frame = tk.Frame(type_frame)
                field_frame.pack(anchor=tk.W, pady=5)
                tk.Label(field_frame, text="Field:").pack(side=tk.LEFT)
                field_entry = tk.Entry(field_frame, width=20)
                field_entry.pack(side=tk.LEFT, padx=5)
                
                # Custom sort pattern entry
                custom_frame = tk.Frame(type_frame)
                custom_frame.pack(anchor=tk.W, pady=5)
                tk.Label(custom_frame, text="Pattern:").pack(side=tk.LEFT)
                pattern_entry = tk.Entry(custom_frame, width=30)
                pattern_entry.pack(side=tk.LEFT, padx=5)
                
                # Custom sort options
                custom_options_frame = tk.Frame(type_frame)
                custom_options_frame.pack(anchor=tk.W, pady=5)
                
                # Sort by option for custom
                custom_sort_by_var = tk.StringVar(value="key")
                tk.Radiobutton(custom_options_frame, text="Apply to Keys", variable=custom_sort_by_var, value="key").pack(side=tk.LEFT, padx=5)
                tk.Radiobutton(custom_options_frame, text="Apply to Values", variable=custom_sort_by_var, value="value").pack(side=tk.LEFT, padx=5)
                tk.Radiobutton(custom_options_frame, text="Apply to Field", variable=custom_sort_by_var, value="field").pack(side=tk.LEFT, padx=5)
                
                # Field entry for custom field sorting
                custom_field_frame = tk.Frame(type_frame)
                custom_field_frame.pack(anchor=tk.W, pady=2)
                tk.Label(custom_field_frame, text="Field for Custom:").pack(side=tk.LEFT)
                custom_field_entry = tk.Entry(custom_field_frame, width=20)
                custom_field_entry.pack(side=tk.LEFT, padx=5)
                
                # Function to update entry states based on radio button selections
                def update_entry_states(*args):
                    sort_type = sort_type_var.get()
                    custom_sort_by = custom_sort_by_var.get()
                    
                    # Enable/disable field entry based on sort type
                    if sort_type == "field":
                        field_entry.config(state=tk.NORMAL)
                    else:
                        field_entry.config(state=tk.DISABLED)
                    
                    # Enable/disable custom pattern entry and options based on sort type
                    if sort_type == "custom":
                        pattern_entry.config(state=tk.NORMAL)
                        # Enable custom sort by radio buttons
                        for widget in custom_options_frame.winfo_children():
                            if isinstance(widget, tk.Radiobutton):
                                widget.config(state=tk.NORMAL)
                        # Enable/disable custom field entry based on custom_sort_by
                        if custom_sort_by == "field":
                            custom_field_entry.config(state=tk.NORMAL)
                        else:
                            custom_field_entry.config(state=tk.DISABLED)
                    else:
                        pattern_entry.config(state=tk.DISABLED)
                        # Disable custom sort by radio buttons
                        for widget in custom_options_frame.winfo_children():
                            if isinstance(widget, tk.Radiobutton):
                                widget.config(state=tk.DISABLED)
                        custom_field_entry.config(state=tk.DISABLED)
                
                # Trace changes to radio button variables
                sort_type_var.trace('w', update_entry_states)
                custom_sort_by_var.trace('w', update_entry_states)
                
                # Initialize entry states
                update_entry_states()
                
                # Helper text for custom patterns
                tk.Label(type_frame, text="Examples: '\\d+' (numbers), '\\d{4}' (4-digit numbers), '[A-Z]+' (letters)", 
                        font=('Arial', 8), fg='gray').pack(anchor=tk.W, padx=20, pady=(5, 0))
                
                # Direction selection
                direction_frame = tk.LabelFrame(layer_dialog, text="Sort Direction", padx=10, pady=10)
                direction_frame.pack(fill=tk.X, padx=10, pady=10)
                
                direction_var = tk.StringVar(value="ascending")
                tk.Radiobutton(direction_frame, text="Ascending (A-Z)", variable=direction_var, value="ascending").pack(anchor=tk.W)
                tk.Radiobutton(direction_frame, text="Descending (Z-A)", variable=direction_var, value="descending").pack(anchor=tk.W)
                
                # Buttons
                button_frame = tk.Frame(layer_dialog)
                button_frame.pack(fill=tk.X, padx=10, pady=10)
                
                def add_layer():
                    sort_type = sort_type_var.get()
                    direction = direction_var.get()
                    field = field_entry.get().strip() if sort_type == "field" else ""
                    pattern = pattern_entry.get().strip() if sort_type == "custom" else ""
                    custom_sort_by = custom_sort_by_var.get() if sort_type == "custom" else ""
                    custom_field = custom_field_entry.get().strip() if sort_type == "custom" and custom_sort_by == "field" else ""
                    
                    if sort_type == "field" and not field:
                        tk.messagebox.showerror("Field Required", "Please enter a field name for field sorting.")
                        return
                    
                    if sort_type == "custom":
                        if not pattern:
                            tk.messagebox.showerror("Pattern Required", "Please enter a regex pattern for custom sorting.")
                            return
                        
                        # Validate regex pattern
                        try:
                            import re
                            re.compile(pattern)
                        except re.error as e:
                            tk.messagebox.showerror("Invalid Pattern", f"Invalid regex pattern: {str(e)}")
                            return
                        
                        if custom_sort_by == "field" and not custom_field:
                            tk.messagebox.showerror("Field Required", "Please enter a field name for custom field sorting.")
                            return
                    
                    # Validate field exists if it's a field sort
                    if sort_type == "field":
                        has_field = False
                        for value in current_data.values():
                            field_value = get_nested_value(value, field)
                            if field_value is not None:
                                has_field = True
                                break
                        
                        if not has_field:
                            tk.messagebox.showerror("Field Not Found", f"Field '{field}' not found in any values.")
                            return
                    
                    # Validate custom field exists if it's a custom field sort
                    if sort_type == "custom" and custom_sort_by == "field":
                        has_field = False
                        for value in current_data.values():
                            field_value = get_nested_value(value, custom_field)
                            if field_value is not None:
                                has_field = True
                                break
                        
                        if not has_field:
                            tk.messagebox.showerror("Field Not Found", f"Field '{custom_field}' not found in any values.")
                            return
                    
                    # Add the layer
                    sort_layers.append({
                        'type': sort_type,
                        'reverse': direction == "descending",
                        'field': field,
                        'pattern': pattern,
                        'custom_sort_by': custom_sort_by,
                        'custom_field': custom_field
                    })
                    
                    update_layers_display()
                    layer_dialog.destroy()
                
                tk.Button(button_frame, text="Add Layer", command=add_layer).pack(side=tk.LEFT, padx=5)
                tk.Button(button_frame, text="Cancel", command=layer_dialog.destroy).pack(side=tk.LEFT, padx=5)
            
            def remove_sort_layer():
                """Remove selected sort layer"""
                selection = layers_listbox.curselection()
                if not selection:
                    tk.messagebox.showwarning("No Selection", "Please select a layer to remove.")
                    return
                
                index = selection[0]
                sort_layers.pop(index)
                update_layers_display()
            
            def move_layer_up():
                """Move selected layer up"""
                selection = layers_listbox.curselection()
                if not selection or selection[0] == 0:
                    return
                
                index = selection[0]
                sort_layers[index], sort_layers[index-1] = sort_layers[index-1], sort_layers[index]
                update_layers_display()
                layers_listbox.selection_set(index-1)
            
            def move_layer_down():
                """Move selected layer down"""
                selection = layers_listbox.curselection()
                if not selection or selection[0] == len(sort_layers) - 1:
                    return
                
                index = selection[0]
                sort_layers[index], sort_layers[index+1] = sort_layers[index+1], sort_layers[index]
                update_layers_display()
                layers_listbox.selection_set(index+1)
            
            def apply_layered_sort():
                """Apply all sort layers in order"""
                if not sort_layers:
                    tk.messagebox.showwarning("No Layers", "Please add at least one sort layer.")
                    return
                
                try:
                    # Start with the original data
                    sorted_data = current_data.copy()
                    
                    # Apply each sort layer in reverse order (Python's sorted is stable)
                    for layer in reversed(sort_layers):
                        if layer['type'] == "key":
                            sorted_data = dict(sorted(sorted_data.items(), 
                                                   key=lambda item: str(item[0]).lower(),
                                                   reverse=layer['reverse']))
                        elif layer['type'] == "value":
                            sorted_data = dict(sorted(sorted_data.items(), 
                                                   key=lambda item: str(item[1]).lower(),
                                                   reverse=layer['reverse']))
                        elif layer['type'] == "field":
                            def sort_key(item):
                                field_value = get_nested_value(item[1], layer['field'])
                                if field_value is None:
                                    return (1, "")  # Put None values at the end
                                return (0, str(field_value).lower())
                            
                            sorted_data = dict(sorted(sorted_data.items(), 
                                                   key=sort_key,
                                                   reverse=layer['reverse']))
                        elif layer['type'] == "custom":
                            import re
                            pattern = layer['pattern']
                            
                            def custom_sort_key(item):
                                # Determine what to sort based on custom_sort_by
                                if layer['custom_sort_by'] == "key":
                                    text_to_sort = str(item[0])
                                elif layer['custom_sort_by'] == "value":
                                    text_to_sort = str(item[1])
                                elif layer['custom_sort_by'] == "field":
                                    field_value = get_nested_value(item[1], layer['custom_field'])
                                    if field_value is None:
                                        return (1, "")  # Put None values at the end
                                    text_to_sort = str(field_value)
                                else:
                                    text_to_sort = str(item[0])  # Default to key
                                
                                # Find matches for the pattern
                                matches = re.findall(pattern, text_to_sort)
                                if matches:
                                    # Convert matches to sortable format
                                    # For numbers, convert to int if possible
                                    sort_values = []
                                    for match in matches:
                                        try:
                                            # Try to convert to int for proper numeric sorting
                                            sort_values.append(int(match))
                                        except ValueError:
                                            try:
                                                # Try to convert to float
                                                sort_values.append(float(match))
                                            except ValueError:
                                                # Keep as string
                                                sort_values.append(match)
                                    return (0, sort_values)
                                else:
                                    return (1, "")  # Put items without matches at the end
                            
                            sorted_data = dict(sorted(sorted_data.items(), 
                                                   key=custom_sort_key,
                                                   reverse=layer['reverse']))
                    
                    # Update the data
                    self.data_sources[data_source]['data'] = sorted_data
                    
                    # Update tab text to show checkmark
                    self.add_checkmark_to_tab(data_source)
                    
                    # Refresh the display
                    refresh_display()
                    
                    # Close the dialog
                    layered_sort_dialog.destroy()
                    
                    tk.messagebox.showinfo("Layered Sort Complete", 
                                         f"Successfully applied {len(sort_layers)} sort layers to {len(sorted_data)} entries.")
                    
                except Exception as e:
                    tk.messagebox.showerror("Sort Error", 
                                          f"An error occurred while applying layered sort:\n{str(e)}")
            
            # Control buttons for layers
            layer_controls = tk.Frame(layers_frame)
            layer_controls.pack(fill=tk.X, pady=(0, 10))
            
            tk.Button(layer_controls, text="Add Layer", command=add_sort_layer).pack(side=tk.LEFT, padx=2)
            tk.Button(layer_controls, text="Remove Layer", command=remove_sort_layer).pack(side=tk.LEFT, padx=2)
            tk.Button(layer_controls, text="Move Up", command=move_layer_up).pack(side=tk.LEFT, padx=2)
            tk.Button(layer_controls, text="Move Down", command=move_layer_down).pack(side=tk.LEFT, padx=2)
            
            # Main dialog buttons
            main_buttons = tk.Frame(main_frame)
            main_buttons.pack(fill=tk.X, pady=(10, 0))
            
            tk.Button(main_buttons, text="Apply Layered Sort", command=apply_layered_sort).pack(side=tk.LEFT, padx=5)
            tk.Button(main_buttons, text="Cancel", command=layered_sort_dialog.destroy).pack(side=tk.LEFT, padx=5)

        def save_to_json():
            """Save the current data to a JSON file"""
            current_data = self.data_sources[data_source]['data']
            if not current_data:
                tk.messagebox.showwarning("No Data", f"No {name} data available to save.")
                return
            
            # Ask user for file location
            file_path = filedialog.asksaveasfilename(
                title=f"Save {name} Data",
                defaultextension=".json",
                filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
                initialfile=f"{name.lower().replace(' ', '_')}_data.json"
            )
            
            if not file_path:
                return  # User cancelled
            
            try:
                # Save data to JSON file with nice formatting
                with open(file_path, 'w', encoding='utf-8') as f:
                    json.dump(current_data, f, indent=2, ensure_ascii=False)
                
                tk.messagebox.showinfo("Save Complete", 
                                     f"Successfully saved {len(current_data)} entries to:\n{file_path}")
                
            except Exception as e:
                tk.messagebox.showerror("Save Error", 
                                      f"An error occurred while saving to JSON:\n{str(e)}")

        def load_from_json():
            """Load data from a JSON file"""
            # Ask user for file location
            file_path = filedialog.askopenfilename(
                title=f"Load {name} Data",
                filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
            )
            
            if not file_path:
                return  # User cancelled
            
            try:
                # Load data from JSON file
                with open(file_path, 'r', encoding='utf-8') as f:
                    loaded_data = json.load(f)
                
                # Validate that it's a dictionary
                if not isinstance(loaded_data, dict):
                    tk.messagebox.showerror("Invalid Format", 
                                          "JSON file must contain a JSON object (dictionary).")
                    return
                
                # Confirm overwrite
                current_data = self.data_sources[data_source]['data']
                if current_data:
                    result = tk.messagebox.askyesno("Confirm Overwrite", 
                                                  f"This will overwrite all current {name} data.\n\n"
                                                  f"Current entries: {len(current_data)}\n"
                                                  f"New entries: {len(loaded_data)}\n\n"
                                                  "Do you want to continue?")
                    if not result:
                        return
                
                # Update the data
                self.data_sources[data_source]['data'] = loaded_data
                
                # Update coordinates combo box with new data
                self.update_coordinates_combo_box(data_source)
                
                # Update tab text to show checkmark
                self.add_checkmark_to_tab(data_source)
                
                # Show success message
                tk.messagebox.showinfo("Load Complete", 
                                     f"Successfully loaded {len(loaded_data)} entries from:\n{file_path}")
                
                # Refresh the display
                refresh_display()
                
            except json.JSONDecodeError as e:
                tk.messagebox.showerror("Invalid JSON", 
                                      f"File is not valid JSON:\n{str(e)}")
            except Exception as e:
                tk.messagebox.showerror("Load Error", 
                                      f"An error occurred while loading from JSON:\n{str(e)}")

        # Add Split Keys button
        split_button = tk.Button(button_frame, text="Split Keys", command=split_keys)
        split_button.pack(side=tk.LEFT, padx=5)
        
        # Add Paste from Clipboard button
        paste_button = tk.Button(button_frame, text="Paste from Clipboard", command=paste_from_clipboard)
        paste_button.pack(side=tk.LEFT, padx=5)
        
        # Add Copy to Clipboard button
        copy_button = tk.Button(button_frame, text="Copy JSON to Clipboard", command=copy_to_clipboard)
        copy_button.pack(side=tk.LEFT, padx=5)
        
        # Add Copy as Table button
        copy_table_button = tk.Button(button_frame, text="Copy as Table", command=copy_as_table)
        copy_table_button.pack(side=tk.LEFT, padx=5)
        
        # Add Modify Keys button
        modify_keys_button = tk.Button(button_frame, text="Modify Keys", command=lambda: self.update_data_source_keys(data_source, parent=view_window, refresh_callback=refresh_display))
        modify_keys_button.pack(side=tk.LEFT, padx=5)
        
        # Add Sort Data button
        sort_button = tk.Button(button_frame, text="Sort Data", command=sort_data)
        sort_button.pack(side=tk.LEFT, padx=5)
        
        # Add Layered Sort Data button
        layered_sort_button = tk.Button(button_frame, text="Layered Sort", command=layered_sort_data)
        layered_sort_button.pack(side=tk.LEFT, padx=5)
        
        # Add Save to JSON button
        save_json_button = tk.Button(button_frame, text="Save to JSON", command=save_to_json)
        save_json_button.pack(side=tk.LEFT, padx=5)
        
        # Add Load from JSON button
        load_json_button = tk.Button(button_frame, text="Load from JSON", command=load_from_json)
        load_json_button.pack(side=tk.LEFT, padx=5)

        # Create a scrolled text widget to display the data
        scrolled_text = scrolledtext.ScrolledText(view_window, width=40, height=20)
        scrolled_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)  # Fill and expand to fill the window

        # Display the data content in the scrolled text widget
        if data:
            for key, value in data.items():
                scrolled_text.insert(tk.END, f"{key}: {value}\n\n")
        else:
            scrolled_text.insert(tk.END, f"No {name} data available.")

        # Keep it editable - removed state='disabled'
        
        def save_edited_changes():
            """Save edited content from the text widget back to the data source"""
            try:
                # Get the content from the text widget
                content = scrolled_text.get("1.0", tk.END).strip()
                
                if not content or content == f"No {name} data available.":
                    tk.messagebox.showwarning("No Data", "No data to save.")
                    return
                
                # Parse the content back into a dictionary
                # Format: "key: value\n\n"
                new_data = {}
                lines = content.split('\n')
                current_key = None
                current_value = []
                
                for line in lines:
                    line = line.strip()
                    if not line:
                        # Empty line means end of current entry
                        if current_key is not None:
                            # Join accumulated value parts
                            value_str = '\n'.join(current_value).strip()
                            # Try to parse as JSON if it looks like JSON, otherwise keep as string
                            try:
                                # Check if it looks like JSON (starts with { or [)
                                if value_str.startswith('{') or value_str.startswith('['):
                                    new_data[current_key] = json.loads(value_str)
                                else:
                                    new_data[current_key] = value_str
                            except:
                                new_data[current_key] = value_str
                            current_key = None
                            current_value = []
                        continue
                    
                    # Check if this line contains a colon (key: value format)
                    if ':' in line and current_key is None:
                        parts = line.split(':', 1)
                        if len(parts) == 2:
                            current_key = parts[0].strip()
                            value_part = parts[1].strip()
                            if value_part:
                                current_value = [value_part]
                            else:
                                current_value = []
                        else:
                            # No colon found, treat as continuation of previous value
                            if current_key is not None:
                                current_value.append(line)
                            else:
                                # First line without colon, treat as key with empty value
                                current_key = line
                                current_value = []
                    else:
                        # Continuation of value
                        if current_key is not None:
                            current_value.append(line)
                        else:
                            # No key yet, treat as key
                            current_key = line
                            current_value = []
                
                # Handle last entry if there's no trailing empty line
                if current_key is not None:
                    value_str = '\n'.join(current_value).strip()
                    try:
                        if value_str.startswith('{') or value_str.startswith('['):
                            new_data[current_key] = json.loads(value_str)
                        else:
                            new_data[current_key] = value_str
                    except:
                        new_data[current_key] = value_str
                
                if not new_data:
                    tk.messagebox.showwarning("No Data", "Could not parse any data from the edited content.")
                    return
                
                # Confirm save
                result = tk.messagebox.askyesno("Confirm Save", 
                                              f"This will overwrite the current {name} data with the edited content.\n\n"
                                              f"Found {len(new_data)} entries in edited content.\n\n"
                                              "Do you want to continue?")
                
                if result:
                    # Update the data
                    self.data_sources[data_source]['data'] = new_data
                    
                    # Update coordinates combo box with new data
                    self.update_coordinates_combo_box(data_source)
                    
                    # Update tab text to show checkmark
                    self.add_checkmark_to_tab(data_source)
                    
                    # Show success message
                    tk.messagebox.showinfo("Save Complete", 
                                         f"Successfully saved {len(new_data)} entries from edited content.")
                    
                    # Refresh the display to show the saved data
                    refresh_display()
                    
            except Exception as e:
                tk.messagebox.showerror("Save Error", 
                                      f"An error occurred while saving edited content:\n{str(e)}")

        # Add Save Changes button to button frame
        save_changes_button = tk.Button(button_frame, text="Save Changes", command=save_edited_changes)
        save_changes_button.pack(side=tk.LEFT, padx=5)
        
        # Configure tags for search highlighting
        scrolled_text.tag_configure("highlight", background="yellow", foreground="black")
        scrolled_text.tag_configure("current_highlight", background="orange", foreground="black")
        
        # Search state variables
        search_matches = []
        current_match_index = [0]  # Using list to make it mutable in nested functions
        
        def clear_highlights():
            """Clear all search highlights"""
            scrolled_text.tag_remove("highlight", "1.0", tk.END)
            scrolled_text.tag_remove("current_highlight", "1.0", tk.END)
        
        def highlight_matches(search_text):
            """Highlight all occurrences of search_text"""
            nonlocal search_matches
            search_matches = []
            current_match_index[0] = 0
            
            # Clear previous highlights
            clear_highlights()
            
            if not search_text:
                match_label.config(text="")
                return
            
            # Find all matches (case-insensitive)
            search_text_lower = search_text.lower()
            content = scrolled_text.get("1.0", tk.END).lower()
            
            # Find all match positions
            start_pos = 0
            while True:
                pos = content.find(search_text_lower, start_pos)
                if pos == -1:
                    break
                search_matches.append(pos)
                start_pos = pos + 1
            
            # Highlight all matches
            for match_pos in search_matches:
                # Convert character position to tkinter index
                idx = f"1.0 + {match_pos} chars"
                end_idx = f"{idx} + {len(search_text)} chars"
                scrolled_text.tag_add("highlight", idx, end_idx)
            
            # Update match counter
            if search_matches:
                match_label.config(text=f"Match 1 of {len(search_matches)}")
                # Highlight the first match differently
                idx = f"1.0 + {search_matches[0]} chars"
                end_idx = f"{idx} + {len(search_text)} chars"
                scrolled_text.tag_add("current_highlight", idx, end_idx)
                # Scroll to first match
                scrolled_text.see(idx)
            else:
                match_label.config(text="No matches found")
        
        def on_search_change(*args):
            """Called when search text changes"""
            search_text = search_var.get()
            highlight_matches(search_text)
        
        def goto_next_match():
            """Navigate to the next match"""
            if not search_matches:
                return
            
            search_text = search_var.get()
            if not search_text:
                return
            
            # Move to next match
            current_match_index[0] = (current_match_index[0] + 1) % len(search_matches)
            
            # Clear current highlight
            scrolled_text.tag_remove("current_highlight", "1.0", tk.END)
            
            # Highlight current match
            match_pos = search_matches[current_match_index[0]]
            idx = f"1.0 + {match_pos} chars"
            end_idx = f"{idx} + {len(search_text)} chars"
            scrolled_text.tag_add("current_highlight", idx, end_idx)
            
            # Scroll to match
            scrolled_text.see(idx)
            
            # Update counter
            match_label.config(text=f"Match {current_match_index[0] + 1} of {len(search_matches)}")
        
        def goto_prev_match():
            """Navigate to the previous match"""
            if not search_matches:
                return
            
            search_text = search_var.get()
            if not search_text:
                return
            
            # Move to previous match
            current_match_index[0] = (current_match_index[0] - 1) % len(search_matches)
            
            # Clear current highlight
            scrolled_text.tag_remove("current_highlight", "1.0", tk.END)
            
            # Highlight current match
            match_pos = search_matches[current_match_index[0]]
            idx = f"1.0 + {match_pos} chars"
            end_idx = f"{idx} + {len(search_text)} chars"
            scrolled_text.tag_add("current_highlight", idx, end_idx)
            
            # Scroll to match
            scrolled_text.see(idx)
            
            # Update counter
            match_label.config(text=f"Match {current_match_index[0] + 1} of {len(search_matches)}")
        
        # Bind search variable to trigger highlighting
        search_var.trace("w", on_search_change)
        
        # Bind navigation buttons
        next_button.config(command=goto_next_match)
        prev_button.config(command=goto_prev_match)
        
        # Bind Enter key to go to next match
        search_entry.bind("<Return>", lambda e: goto_next_match())
        search_entry.bind("<Shift-Return>", lambda e: goto_prev_match())
        
        # Set the reapply search function for refresh_display
        def reapply_search_func():
            search_text = search_var.get()
            if search_text:
                highlight_matches(search_text)
        reapply_search[0] = reapply_search_func
        
        # Trigger initial search highlighting if search term was provided
        if initial_search:
            highlight_matches(initial_search)



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
        if self.excel_mgr.wb:
            self.excel_mgr.save_workbook()
        self.excel_mgr.close_workbook()

    def release_excel_connection(self):
        """Release the xlwings connection to the current destination's workbook, allowing user to save independently"""
        try:
            current_destination = self.get_current_destination()
            if current_destination:
                dest_path, _, _, _, _, _ = self.get_destination_ui_values(current_destination)
                if dest_path and self.excel_mgr.wb:
                    # Check if the current workbook matches this destination
                    try:
                        current_wb_path = os.path.normpath(os.path.abspath(self.excel_mgr.wb.fullname)).lower()
                        dest_path_normalized = os.path.normpath(os.path.abspath(dest_path)).lower()
                        if current_wb_path == dest_path_normalized:
                            # Release the connection but keep Excel open
                            self.excel_mgr.release_connection()
                            # Remove from cache so it won't be switched back to
                            if self.excel_mgr.original_path:
                                normalized_path = self.excel_mgr._normalize_path(self.excel_mgr.original_path)
                                if normalized_path in self.excel_mgr.open_workbooks:
                                    del self.excel_mgr.open_workbooks[normalized_path]
                            tk.messagebox.showinfo("Excel Released", 
                                                 f"xlwings connection to '{current_destination}' workbook has been released. The workbook remains open in Excel and you can now save it independently.")
                        else:
                            tk.messagebox.showinfo("Info", 
                                                 f"Current workbook does not match destination '{current_destination}'. No connection to release.")
                    except Exception as e:
                        print(f"Error checking workbook path: {e}")
                        # Try to release anyway
                        if self.excel_mgr.wb or self.excel_mgr.app:
                            self.excel_mgr.release_connection()
                            tk.messagebox.showinfo("Excel Released", 
                                                 "xlwings connection has been released. The workbook remains open in Excel.")
                else:
                    tk.messagebox.showinfo("Info", 
                                         f"No workbook is currently open for destination '{current_destination}'.")
            else:
                # No destination selected, try to release current workbook if any
                if self.excel_mgr.wb or self.excel_mgr.app:
                    self.excel_mgr.release_connection()
                    tk.messagebox.showinfo("Excel Released", 
                                         "xlwings connection has been released. The workbook remains open in Excel and you can now save it independently.")
                else:
                    tk.messagebox.showinfo("Info", "No Excel connection to release.")

        except Exception as e:
            tk.messagebox.showerror("Error", 
                                  f"Error releasing Excel connection: {str(e)}")
    
    def release_all_excel_connections(self):
        """Release all xlwings connections to all open workbooks, allowing user to save independently"""
        try:
            if not self.excel_mgr.open_workbooks:
                tk.messagebox.showinfo("Info", "No Excel workbooks are currently open.")
                return
            
            released_count = 0
            for normalized_path, workbook_info in list(self.excel_mgr.open_workbooks.items()):
                try:
                    # Set current state to this workbook
                    self.excel_mgr.wb = workbook_info['wb']
                    self.excel_mgr.app = workbook_info['app']
                    self.excel_mgr.is_dirty = workbook_info['is_dirty']
                    self.excel_mgr.original_path = workbook_info['original_path']
                    self.excel_mgr.temp_path = workbook_info['temp_path']
                    # Release connection (this will set wb and app to None)
                    self.excel_mgr.release_connection()
                    released_count += 1
                except Exception as e:
                    print(f"Error releasing workbook {workbook_info.get('original_path', 'unknown')}: {e}")
            
            # Clear the cache and reset current references
            self.excel_mgr.open_workbooks.clear()
            self.excel_mgr.wb = None
            self.excel_mgr.app = None
            self.excel_mgr.is_dirty = False
            self.excel_mgr.original_path = None
            self.excel_mgr.temp_path = None
            
            tk.messagebox.showinfo("Excel Released", 
                                 f"Released xlwings connections to {released_count} workbook(s). The workbooks remain open in Excel and you can now save them independently.")

        except Exception as e:
            tk.messagebox.showerror("Error", 
                                  f"Error releasing Excel connections: {str(e)}")

    def show_generation_report(self):
        """Show a detailed report of the datasheet generation process"""
        try:
            # Count the number of sheets created/updated
            num_sheets = len(self.new_sheets) if self.new_sheets else 0
            
            # Get additional information for the report
            # Use the actual source sheet name that was used in the process
            source_sheet = getattr(self, 'last_used_source_sheet', "Unknown")
            datasheet_prefix = getattr(self, 'last_used_datasheet_prefix', "Unknown")
            rows_per_sheet = getattr(self, 'last_used_rows_per_sheet', "Unknown")
            
            # Get statistics from the generation process
            green_highlighted = 0
            updated_cells = 0
            if hasattr(self, 'generation_statistics') and self.generation_statistics:
                green_highlighted = self.generation_statistics.get('green_highlighted_cells', 0)
                updated_cells = self.generation_statistics.get('updated_cells', 0)
            
            # Create the report message
            report_message = f"""Datasheet Generation Complete!

📊 Generation Summary:
• Source Sheet: {source_sheet}
• Datasheet Prefix: {datasheet_prefix}
• Sheets Created/Updated: {num_sheets}
• Rows per Sheet: {rows_per_sheet}

📈 Cell Processing Statistics:
• Cells Highlighted Green: {green_highlighted}
• Cells Updated: {updated_cells}

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
        if not self.destination_datasheet:
            print("DEBUG: No datasheets file set, returning empty list")
            return []
        
        # Check if we have cached sheet names and they're still valid
        if (not force_reload and 
            hasattr(self, '_cached_sheet_names') and 
            hasattr(self, '_cached_sheet_file') and 
            self._cached_sheet_file == self.destination_datasheet):
            print("DEBUG: Using cached sheet names")
            return self._cached_sheet_names
        
        try:
            print(f"DEBUG: Loading workbook: {self.destination_datasheet}")
            # Use read_only=True and data_only=True for faster loading
            wb = openpyxl.load_workbook(self.destination_datasheet, read_only=True, data_only=True)
            print("DEBUG: Workbook loaded successfully")
            sheet_names = wb.sheetnames
            print(f"DEBUG: Found {len(sheet_names)} sheets: {sheet_names}")
            wb.close()  # Explicitly close the workbook
            print("DEBUG: Workbook closed")
            
            # Cache the sheet names
            self._cached_sheet_names = sheet_names
            self._cached_sheet_file = self.destination_datasheet
            print("DEBUG: Sheet names cached")
            return sheet_names
        except Exception as e:
            print(f"DEBUG: Error getting sheet names from {self.destination_datasheet}: {e}")
            return []

    def on_closing(self):
        # Stop Excel monitoring
        self.stop_excel_monitoring()
        
        # Close all open workbooks when closing the app
        try:
            self.excel_mgr.cleanup(close_all=True)
        except Exception as e:
            print(f"Warning: Error closing workbooks: {e}")
        
        self.root.destroy()


    # endregion

    # region UI Utilities





    def browse_coordinate_value(self, entry):
        """Browse for coordinate value file"""
        filename = filedialog.askopenfilename(
            filetypes=[("All supported files", "*.json;*.xlsx;*.xls;*.xlsm"), 
                      ("JSON files", "*.json"), 
                      ("Excel files", "*.xlsx;*.xls;*.xlsm"),
                      ("All files", "*.*")]
        )
        
        if filename:
            entry.delete(0, tk.END)
            entry.insert(0, filename)
            self.coordinate_value_path = filename

    def browse_datasheets(self, entry, destination=None):
        """Browse for datasheets file"""
        print("DEBUG: Starting browse_datasheets...")
        filename = filedialog.askopenfilename(
            filetypes=[("All supported files", "*.json;*.xlsx;*.xls;*.xlsm"), 
                      ("JSON files", "*.json"), 
                      ("Excel files", "*.xlsx;*.xls;*.xlsm"),
                      ("All files", "*.*")]
        )
        
        if filename:
            try:
                print("DEBUG: File selected, updating entry...")
                entry.delete(0, tk.END)
                entry.insert(0, filename)
                
                # If destination is provided, update that destination's path
                if destination and destination in self.destinations:
                    self.destinations[destination]['path'] = filename
                    print(f"DEBUG: Destination '{destination}' path set to: {filename}")
                else:
                    # Legacy: update global destination_datasheet
                    self.destination_datasheet = filename
                    print(f"DEBUG: Destination Datasheet set to: {self.destination_datasheet}")
                
                self._cached_sheet_names = None
                self._cached_sheet_file = None
                print("DEBUG: Cleared cached sheet names for new file")
                    
            except Exception as e:
                print(f"DEBUG: Error setting datasheets file: {e}")
                # Still set the filename even if there's an error
                if destination and destination in self.destinations:
                    self.destinations[destination]['path'] = filename
                else:
                    self.destination_datasheet = filename

    def browse_data_source(self, entry, data_source):
        """Browse for a file for a specific data source"""
        filename = filedialog.askopenfilename(
            filetypes=[("All supported files", "*.json;*.xlsx;*.xls;*.xlsm"), 
                      ("JSON files", "*.json"), 
                      ("Excel files", "*.xlsx;*.xls;*.xlsm"),
                      ("All files", "*.*")]
        )
        
        if filename:
            entry.delete(0, tk.END)
            entry.insert(0, filename)
            self.data_sources[data_source]['path'] = filename
            print(f"Set {data_source} path to {filename}")

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
        
        # Find the data source by name
        data_source = text  # Direct reference since text is the data source name
        if data_source:
            self.configure_data_source(data_source)
        else:
            print(f"Unknown configuration type: {text}")
    
    
    def ensure_global_entries_exist(self):
        """Ensure that global entries exist in self.entries list using stored references"""
        # Check if we have the required global entries
        required_entries = ['datasheet_coord', 'ds_str', 'rows_per_sheet', 'sig_figs', 'rounding_tolerance']
        existing_variables = [var for _, var in self.entries]
        
        missing_entries = [var for var in required_entries if var not in existing_variables]
        
        if missing_entries:
            print(f'DEBUG: Missing global entries: {missing_entries}')
            print(f'DEBUG: Current entries: {existing_variables}')
            # Add missing entries using stored references
            self.add_global_entries_from_references(missing_entries)
        else:
            print(f'DEBUG: All required global entries found: {required_entries}')
    
    def add_global_entries_from_references(self, missing_entries):
        """Add missing global entries using stored references"""
        print(f'DEBUG: Adding missing global entries from references: {missing_entries}')
        
        # Map of variable names to stored entry references
        entry_references = {
            'datasheet_coord': getattr(self, 'datasheet_coord_entry', None),
            'ds_str': getattr(self, 'ds_str_entry', None),
            'rows_per_sheet': getattr(self, 'rows_per_sheet_entry', None),
            'sig_figs': getattr(self, 'sig_figs_entry', None),
            'rounding_tolerance': getattr(self, 'rounding_tolerance_entry', None)
        }
        
        # Add each missing entry using stored reference
        for variable in missing_entries:
            entry_widget = entry_references.get(variable)
            
            if entry_widget:
                # Check if this entry is already in self.entries
                already_exists = any(var == variable for _, var in self.entries)
                if not already_exists:
                    self.entries.append((entry_widget, variable))
                    print(f"DEBUG: Added global entry for {variable} from stored reference")
                else:
                    print(f"DEBUG: Global entry for {variable} already exists")
            else:
                print(f"ERROR: No stored reference found for {variable}")
    
    def get_destination_ui_values(self, destination):
        """Get current values directly from destination-specific UI entries"""
        try:
            if destination not in self.destination_widgets:
                print(f"No UI widgets found for destination {destination}")
                # Fall back to stored destination config
                config = self.destinations.get(destination, {})
                return (
                    config.get('path', ''),
                    config.get('datasheet_coord', ''),
                    config.get('ds_str', ''),
                    config.get('rows_per_sheet', 1),
                    config.get('sig_figs', 4),
                    config.get('rounding_tolerance', 1e-2)
                )
            
            widgets = self.destination_widgets[destination]
            
            # Get values from entry widgets
            path = widgets.get('datasheet_entry', tk.Entry()).get().strip() if 'datasheet_entry' in widgets else ''
            datasheet_coord = widgets.get('datasheet_coord_entry', tk.Entry()).get().strip() if 'datasheet_coord_entry' in widgets else ''
            ds_str = widgets.get('ds_str_entry', tk.Entry()).get().strip() if 'ds_str_entry' in widgets else ''
            
            rows_per_sheet_str = widgets.get('rows_per_sheet_entry', tk.Entry()).get().strip() if 'rows_per_sheet_entry' in widgets else '1'
            rows_per_sheet = int(rows_per_sheet_str) if rows_per_sheet_str.isdigit() else 1
            
            sig_figs_str = widgets.get('sig_figs_entry', tk.Entry()).get().strip() if 'sig_figs_entry' in widgets else '4'
            sig_figs = int(sig_figs_str) if sig_figs_str.isdigit() else 4
            
            tolerance_str = widgets.get('rounding_tolerance_entry', tk.Entry()).get().strip() if 'rounding_tolerance_entry' in widgets else '1e-2'
            try:
                rounding_tolerance = float(tolerance_str) if tolerance_str.replace('.', '').replace('-', '').replace('e', '').replace('E', '').replace('+', '').isdigit() or 'e' in tolerance_str.lower() else 1e-2
            except:
                rounding_tolerance = 1e-2
            
            return (path, datasheet_coord, ds_str, rows_per_sheet, sig_figs, rounding_tolerance)
            
        except Exception as e:
            print(f"Error getting destination UI values for {destination}: {e}")
            # Fall back to stored config
            config = self.destinations.get(destination, {})
            return (
                config.get('path', ''),
                config.get('datasheet_coord', ''),
                config.get('ds_str', ''),
                config.get('rows_per_sheet', 1),
                config.get('sig_figs', 4),
                config.get('rounding_tolerance', 1e-2)
            )

    def get_data_source_ui_values(self, data_source):
        """Get current values directly from data_source specific UI entries using stored references"""
        try:
            # Get the stored entry references from the dictionary
            if not hasattr(self, 'data_source_ui_entries') or data_source not in self.data_source_ui_entries:
                print(f"No UI entries found for {data_source}")
                return None, None, False
            
            entries = self.data_source_ui_entries[data_source]
            source_sheet_entry = entries.get('source_sheet_entry')
            top_tag_entry = entries.get('top_tag_entry')
            partial_match_var = entries.get('partial_match_var')
            
            print(f"DEBUG: Found entries for {data_source}: source_sheet_entry={source_sheet_entry is not None}, top_tag_entry={top_tag_entry is not None}")
            
            # Handle missing UI entries - get from config if available
            if not source_sheet_entry:
                print(f"No source_sheet_entry found for {data_source}")
                source_sheet_name = None
            else:
                source_sheet_name = source_sheet_entry.get().strip()
            
            if not top_tag_entry:
                print(f"No top_tag_entry found for {data_source}, using config value")
                # Fallback to config value if UI entry doesn't exist yet
                top_tag = self.data_sources[data_source].get('top_tag', '') if data_source in self.data_sources else ''
            else:
                top_tag = top_tag_entry.get().strip()
            
            if not partial_match_var:
                print(f"No partial_match_var found for {data_source}, using config value")
                # Fallback to config value if UI entry doesn't exist yet
                partial_match = self.data_sources[data_source].get('partial_match', False) if data_source in self.data_sources else False
            else:
                partial_match = partial_match_var.get()
            
            print(f"DEBUG: Retrieved values for {data_source}: source_sheet='{source_sheet_name}', top_tag='{top_tag}', partial_match={partial_match}")
            
            return source_sheet_name, top_tag, partial_match
            
        except Exception as e:
            print(f"Error getting UI values for data source {data_source}: {e}")
            return None, None, False
    
    def set_data_source_ui_values(self, data_source, source_sheet_name, top_tag, partial_match=False):
        """Set values to data_source specific UI entries"""
        try:
            # Get the stored entry references from the dictionary
            if not hasattr(self, 'data_source_ui_entries') or data_source not in self.data_source_ui_entries:
                print(f"No UI entries found for {data_source}")
                return False
            
            entries = self.data_source_ui_entries[data_source]
            source_sheet_entry = entries.get('source_sheet_entry')
            top_tag_entry = entries.get('top_tag_entry')
            partial_match_var = entries.get('partial_match_var')
            
            if not source_sheet_entry or not top_tag_entry:
                print(f"Incomplete UI entries found for {data_source}")
                return False
            
            # Set values to the entries
            if source_sheet_name:
                source_sheet_entry.set(source_sheet_name)
            if top_tag:
                top_tag_entry.delete(0, tk.END)
                top_tag_entry.insert(0, top_tag)
            if partial_match_var:
                partial_match_var.set(partial_match)
            
            return True
            
        except Exception as e:
            print(f"Error setting UI values for data source {data_source}: {e}")
            return False
    
    def configure_data_source(self, data_source):
        """Configure a specific data source using the centralized system"""
        config = self.data_sources[data_source]
        name = data_source
        
        file_path = self.data_sources[data_source]['path']
        current_headers = self.data_sources[data_source]['headers']
        current_selection = self.data_sources[data_source]['selected_sheets']
        current_tolerance = self.blank_cell_tolerance

        if not file_path or not os.path.exists(file_path):
            messagebox.showwarning("File Not Found", f"Please select a valid {name} file first.", parent=self.root)
            return

        result = configure_source_data_dialog(self.root, f"Configure {name} Source",
                                              current_headers, file_path, current_selection, current_tolerance)

        if result:
            self.data_sources[data_source]['headers'] = result["headers"]
            self.data_sources[data_source]['selected_sheets'] = result["selected_sheets"]
            self.blank_cell_tolerance = result["tolerance"]
            print(f"{name} configuration updated.")


    def update_source_sheet_comboboxes_with_destination_sheets(self):
        """Update all Source Sheet Name comboboxes with cached sheet names from destination datasheet"""
        print("DEBUG: Updating source sheet comboboxes...")
        if not hasattr(self, 'data_source_ui_entries'):
            print("DEBUG: No data source UI entries found")
            return
        
        # Get the cached sheet names
        sheet_names = self.get_sheet_names(force_reload=True)
        print(f"DEBUG: Retrieved {len(sheet_names)} sheet names: {sheet_names}")
        
        # Update each data source's combobox
        for data_source, ui_entries in self.data_source_ui_entries.items():
            # Skip data sources that no longer exist (e.g., after new project)
            if data_source not in self.data_sources:
                continue
            if 'source_sheet_entry' in ui_entries and ui_entries['source_sheet_entry']:
                try:
                    source_sheet_entry = ui_entries['source_sheet_entry']
                    source_sheet_entry['values'] = sheet_names
                    print(f"DEBUG: Updated combobox for data source '{data_source}' with {len(sheet_names)} sheets")
                except Exception as e:
                    print(f"DEBUG: Error updating combobox for data source '{data_source}': {e}")

    def configure_ds(self, destination=None):
        print("DEBUG: Starting configure_ds...")
        
        # If no destination provided, use current destination
        if destination is None:
            destination = self.get_current_destination()
            print(f"DEBUG: No destination provided, using current destination: {destination}")
        
        # Get the path from the appropriate entry field
        if destination and destination in self.destination_widgets:
            # Use destination-specific entry
            widgets = self.destination_widgets[destination]
            if 'datasheet_entry' in widgets:
                current_path = widgets['datasheet_entry'].get()
                if current_path:
                    self.destinations[destination]['path'] = current_path
                    self.destination_datasheet = current_path  # Also set global for compatibility
                    print(f"DEBUG: Using destination '{destination}' datasheet path from entry: {current_path}")
                else:
                    messagebox.showwarning("No Path", f"Please specify a datasheet path for destination '{destination}' before mapping coordinates.")
                    return
            else:
                messagebox.showwarning("No Path", f"Destination '{destination}' does not have a datasheet path configured.")
                return
        else:
            # Use global entry (legacy)
            if hasattr(self, 'datasheet_entry'):
                current_path = self.datasheet_entry.get()
                if current_path:
                    self.destination_datasheet = current_path
                    print(f"DEBUG: Using destination datasheet path from entry: {current_path}")
                else:
                    messagebox.showwarning("No Path", "Please specify a datasheet path before mapping coordinates.")
                    return
            else:
                messagebox.showwarning("No Path", "No datasheet path configured.")
                return
        
        # Initialize Excel with the current workbook (will switch to it if already open)
        self.init_excel()
        print("DEBUG: init_excel completed, setting cached sheet names...")
        
        # Force reload and cache the sheet names
        sheet_names = self.get_sheet_names(force_reload=True)
        print(f"DEBUG: Cached {len(sheet_names)} sheet names")
        
        # Update all Source Sheet Name comboboxes
        self.update_source_sheet_comboboxes_with_destination_sheets()
        
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
        data_source = self.current_data_source_for_context
        if not data_source:
            return
        existing_conversion = self.data_sources[data_source]['coordinate_conversions'].get(coord, {})

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

                self.data_sources[data_source]['coordinate_conversions'][coord] = {
                    'in_unit': in_unit_combo.get(),
                    'out_unit': out_unit_combo.get(),
                    'formula': formula_str
                }
                print(f"Saved conversion for {coord}: {self.data_sources[data_source]['coordinate_conversions'][coord]}")
            else:
                # If formula is empty, remove the conversion
                if coord in self.data_sources[data_source]['coordinate_conversions']:
                    del self.data_sources[data_source]['coordinate_conversions'][coord]
                    print(f"Removed conversion for {coord} (empty formula)")

            # Update listboxes in the main window (call the method on the tab)
            # Find the current data source tab and its coordinates tab
            current_tab_id = self.data_sources_notebook.select()
            current_tab_text = self.data_sources_notebook.tab(current_tab_id, "text")
            current_data_source = self.strip_tab_indicators(current_tab_text).strip()
            
            # Find the coordinates tab within the current data source tab
            current_tab = self.data_sources_notebook.nametowidget(current_tab_id)
            for child in current_tab.winfo_children():
                if isinstance(child, ttk.Notebook):
                    for i in range(child.index('end')):
                        if child.tab(i, "text") == 'Coordinates':
                            coordinates_tab = child.winfo_children()[i]
                            if hasattr(coordinates_tab, 'reinitialize'):
                                coordinates_tab.reinitialize()
                            break
                    break

            dialog.destroy()

        ttk.Button(button_frame, text="OK", command=save_conversion).pack(side=RIGHT, padx=5)
        ttk.Button(button_frame, text="Cancel", command=dialog.destroy).pack(side=RIGHT)

        main_frame.columnconfigure(1, weight=1)
        self.root.wait_window(dialog)

    def remove_conversion(self):
        data_source = self.current_data_source_for_context
        if data_source and self.selected_coord_for_context and self.selected_coord_for_context in self.data_sources[data_source]['coordinate_conversions']:
            coord = self.selected_coord_for_context
            del self.data_sources[data_source]['coordinate_conversions'][coord]
            print(f"Removed conversion for {coord}")
            # Update listboxes
            # Find the current data source tab and its coordinates tab
            current_tab_id = self.data_sources_notebook.select()
            current_tab = self.data_sources_notebook.nametowidget(current_tab_id)
            for child in current_tab.winfo_children():
                if isinstance(child, ttk.Notebook):
                    for i in range(child.index('end')):
                        if child.tab(i, "text") == 'Coordinates':
                            coordinates_tab = child.winfo_children()[i]
                            if hasattr(coordinates_tab, 'reinitialize'):
                                coordinates_tab.reinitialize()
                            break
                    break
        else:
            print("No conversion selected or found to remove.")

    def open_combination_dialog(self):
        if not self.selected_coord_for_context:
            return

        coord = self.selected_coord_for_context
        data_source = self.current_data_source_for_context
        if not data_source:
            return
        existing_combination = self.data_sources[data_source]['coordinate_combinations'].get(coord, {})

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
        for data_source in self.get_all_data_sources():
            coord_values = self.data_sources[data_source]['coordinate_values']
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

            data_source = self.current_data_source_for_context
            if not data_source:
                return
            self.data_sources[data_source]['coordinate_combinations'][coord] = {
                'operation': operation,
                'combines': combines
            }
            print(f"Saved combination for {coord}: {self.data_sources[data_source]['coordinate_combinations'][coord]}")

            # Update listboxes
            # Find the current data source tab and its coordinates tab
            current_tab_id = self.data_sources_notebook.select()
            current_tab = self.data_sources_notebook.nametowidget(current_tab_id)
            for child in current_tab.winfo_children():
                if isinstance(child, ttk.Notebook):
                    for i in range(child.index('end')):
                        if child.tab(i, "text") == 'Coordinates':
                            coordinates_tab = child.winfo_children()[i]
                            if hasattr(coordinates_tab, 'reinitialize'):
                                coordinates_tab.reinitialize()
                            break
                    break

            dialog.destroy()

        ttk.Button(dialog_button_frame, text="OK", command=save_combination).pack(side=RIGHT, padx=5)
        ttk.Button(dialog_button_frame, text="Cancel", command=dialog.destroy).pack(side=RIGHT)

        main_frame.columnconfigure(1, weight=1)
        self.root.wait_window(dialog)

    def remove_combination(self):
        data_source = self.current_data_source_for_context
        if data_source and self.selected_coord_for_context and self.selected_coord_for_context in self.data_sources[data_source]['coordinate_combinations']:
            coord = self.selected_coord_for_context
            del self.data_sources[data_source]['coordinate_combinations'][coord]
            print(f"Removed combination for {coord}")
            # Update listboxes
            # Find the current data source tab and its coordinates tab
            current_tab_id = self.data_sources_notebook.select()
            current_tab = self.data_sources_notebook.nametowidget(current_tab_id)
            for child in current_tab.winfo_children():
                if isinstance(child, ttk.Notebook):
                    for i in range(child.index('end')):
                        if child.tab(i, "text") == 'Coordinates':
                            coordinates_tab = child.winfo_children()[i]
                            if hasattr(coordinates_tab, 'reinitialize'):
                                coordinates_tab.reinitialize()
                            break
                    break
        else:
            print("No combination selected or found to remove.")

    def show_excel_processing_dialog(self, data_source, file_path):
        """Show custom dialog for Excel file processing options"""
        config = self.data_sources[data_source]
        name = data_source
        

        # Create dialog window
        dialog = tk.Toplevel(self.root)
        dialog.title("Excel File Processing")
        dialog.transient(self.root)
        dialog.grab_set()
        
        dialog_width = 500
        dialog_height = 400
        dialog.geometry(f"{dialog_width}x{dialog_height}")
        
        # Center the dialog over the main window
        center_window_over_parent(dialog)

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
        
        # Determine which data source this coordinate belongs to
        coord_source = None
        current_source_key = None
        available_keys = []
        
        for data_source in self.get_all_data_sources():
            coordinate_values = self.data_sources[data_source]['coordinate_values']
            if coord in coordinate_values:
                coord_source = data_source
                current_source_key = coordinate_values[coord]
                # Get available keys for this data source
                data = self.data_sources[data_source]['data']
                for key, value in data.items():
                    available_keys = list(value.keys())
                    break
                break
        
        if coord_source is None:
            print(f"Coordinate {coord} not found in any data source coordinate values")
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
        ttk.Label(main_frame, text=f"Type: {coord_source.upper()}").grid(row=1, column=0, sticky=W, pady=2)
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
                # Update coordinate values for the data source
                coord_values = self.data_sources[coord_source]['coordinate_values']
                coord_values[coord] = new_key
                self.data_sources[coord_source]['coordinate_values'] = coord_values
                print(f"Changed source key for {coord} from '{current_source_key}' to '{new_key}'")
                
                # Update listboxes
                # Find the current data source tab and its coordinates tab
                current_tab_index = self.data_sources_notebook.index('current')
                current_tab = self.data_sources_notebook.winfo_children()[current_tab_index]
                for child in current_tab.winfo_children():
                    if isinstance(child, ttk.Notebook):
                        for i in range(child.index('end')):
                            if child.tab(i, "text") == 'Coordinates':
                                coordinates_tab = child.winfo_children()[i]
                                if hasattr(coordinates_tab, 'reinitialize'):
                                    coordinates_tab.reinitialize()
                                break
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
        center_window_over_parent(regex_window)
        ExcelRegexSearchApp(regex_window)
    
    def open_semantic_matcher(self):
        """Opens the Semantic Matcher window."""
        from semantic_matcher import SemanticMatcherApp
        semantic_window = tk.Toplevel(self.root)
        SemanticMatcherApp(semantic_window, self)

    # Semantic similarity methods
    def load_semantic_model_async(self):
        """Load the sentence transformer model in a background thread"""
        print("DEBUG: load_semantic_model_async() called")
        
        def load_model():
            print("DEBUG: load_model() thread started")
            try:
                self.loading_model = True
                print("DEBUG: Set loading_model = True")
                print("DEBUG: Importing sentence_transformers...")
                from sentence_transformers import SentenceTransformer
                print("DEBUG: Import successful, creating model...")
                # Use a lightweight model for faster loading
                self.semantic_model = SentenceTransformer('all-MiniLM-L6-v2')
                print("DEBUG: Model created successfully")
                self.model_loaded = True
                self.loading_model = False
                print("DEBUG: Set model_loaded = True, loading_model = False")
                
                # Update UI in main thread
                print("DEBUG: Scheduling on_semantic_model_loaded callback")
                self.root.after(0, self.on_semantic_model_loaded)
            except Exception as e:
                print(f"DEBUG: Exception occurred in load_model: {type(e).__name__}: {str(e)}")
                import traceback
                traceback.print_exc()
                self.loading_model = False
                self.root.after(0, lambda: self.on_semantic_model_error(str(e)))
        
        print("DEBUG: Creating and starting thread...")
        thread = threading.Thread(target=load_model, daemon=True)
        thread.start()
        print("DEBUG: Thread started")
    
    def on_semantic_model_loaded(self):
        """Called when semantic model is successfully loaded"""
        print("DEBUG: on_semantic_model_loaded() called")
        print("Semantic model loaded successfully!")
        # Update status label if it exists
        if hasattr(self, 'semantic_status_label'):
            print("DEBUG: Updating semantic_status_label to Ready")
            self.semantic_status_label.config(text="Semantic model: Ready", foreground="green")
        else:
            print("DEBUG: WARNING - semantic_status_label does not exist!")
        
        messagebox.showinfo("Success", "Semantic model loaded successfully!")
    
    def on_semantic_model_error(self, error_msg):
        """Called when semantic model loading fails"""
        print(f"DEBUG: on_semantic_model_error() called")
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
    
    def auto_map_coordinate_semantic(self, data_source, current_coord, min_score=0.3):
        """Version of semantic mapping without user feedback dialogs"""
        print(f"\n=== AUTO_MAP_COORDINATE_SEMANTIC DEBUG ===")
        print(f"Input parameters:")
        print(f"  - data_source: {data_source}")
        print(f"  - current_coord: {current_coord}")
        print(f"  - min_score: {min_score}")
        
        try:
            # Get cell values above and left using the existing method
            sheet = xw.apps.active.books.active.sheets.active
            print(f"  - Active sheet: {sheet.name}")
            
            above_value, left_value = self.get_cell_values_above_and_left(sheet, current_coord)
            print(f"  - Above value: '{above_value}'")
            print(f"  - Left value: '{left_value}'")
            
            # Get available options for this data source
            data = self.data_sources[data_source]['data']
            print(f"  - data source data retrieved: {data is not None}")
            if data:
                print(f"  - Data keys count: {len(data)}")
                print(f"  - First data key: {list(data.keys())[0] if data else 'None'}")
            
            if not data:
                print("  - ERROR: No data found for data_source")
                return None, 0.0, None
            
            # Get the options (keys from the first value dict)
            options = []
            for key, value in data.items():
                options = list(value.keys())
                break
            
            print(f"  - Available options: {options}")
            print(f"  - Options count: {len(options)}")
            
            if not options:
                print("  - ERROR: No options found in data")
                return None, 0.0, None
            
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
                return None, 0.0, None
            
            # Check if semantic model is available
            print(f"  - DEBUG: model_loaded = {self.model_loaded}")
            print(f"  - DEBUG: hasattr semantic_model = {hasattr(self, 'semantic_model')}")
            print(f"  - DEBUG: semantic_model value = {self.semantic_model}")
            
            if not hasattr(self, 'semantic_model') or self.semantic_model is None:
                print("  - ERROR: Semantic model not available")
                print("  - HINT: Click 'Load Semantic Model' button first!")
                return None, 0.0, None
            
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
            return None, 0.0, None

    def automap_range(self, data_source, full_selection, combo_box, coord_entry, listbox, min_score=0.3, destination=None):
        """Iterate through a range of cells and perform semantic mapping"""
        try:
            # Get current destination if not provided
            if destination is None:
                destination = self.get_current_destination()
                if destination is None:
                    destination = "Default"
            
            # Parse the range
            clean_selection = full_selection.replace('$', '')
            
            # Extract all cell addresses from the range
            cell_addresses = self.extract_cell_addresses_from_range(clean_selection)
            print('debug: cell_addresses', cell_addresses)
            # Process all cells using the unified processor
            successful_mappings, mapping_results = self.process_cells_for_automap(data_source, cell_addresses, min_score, destination)
            
            # Update the listbox to show all new mappings
            listbox_obj = self.data_source_widgets.get(data_source, {}).get('listbox')
            if listbox_obj:
                self.update_single_listbox(data_source, listbox_obj, destination)
            
            # Update tab color since coordinate maps may have changed
            self.update_tab_color_for_data_source(data_source)
            
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

    def automap_noncontiguous(self, data_source, full_selection, combo_box, coord_entry, listbox, min_score=0.3, destination=None):
        """Process non-contiguous selections (multiple ranges separated by commas)"""
        try:
            # Get current destination if not provided
            if destination is None:
                destination = self.get_current_destination()
                if destination is None:
                    destination = "Default"
            
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
                    successful_mappings, mapping_results = self.process_cells_for_automap(data_source, cell_addresses, min_score, destination)
                else:
                    # This is a single cell (e.g., C3)
                    successful_mappings, mapping_results = self.process_cells_for_automap(data_source, [range_addr], min_score, destination)
                
                total_successful_mappings += successful_mappings
                all_mapping_results.extend(mapping_results)
                print(f"Range {i+1} completed: {successful_mappings} mappings")
            
            # Update the listbox to show all new mappings
            listbox_obj = self.data_source_widgets.get(data_source, {}).get('listbox')
            if listbox_obj:
                self.update_single_listbox(data_source, listbox_obj, destination)
            
            # Update tab color since coordinate maps may have changed
            self.update_tab_color_for_data_source(data_source)
            
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



    def update_single_listbox(self, data_source, listbox, destination=None):
        """Update a single listbox for a specific data source"""
        # Get current destination if not provided
        if destination is None:
            destination = self.get_current_destination()
            if destination is None:
                destination = "Default"
        
        listbox.delete(0, tk.END)
        # Access coordinate_values per destination
        if destination not in self.data_sources[data_source]['coordinate_values']:
            self.data_sources[data_source]['coordinate_values'][destination] = {}
        coordinate_values = self.data_sources[data_source]['coordinate_values'][destination]
        for key, value in coordinate_values.items():
            coord_display = f"{key}: {value}"
            if key in self.data_sources[data_source]['coordinate_conversions']:
                conv = self.data_sources[data_source]['coordinate_conversions'][key]
                coord_display += f" [{conv.get('in_unit', '?')}->{conv.get('out_unit', '?')}]"
            if key in self.data_sources[data_source]['coordinate_combinations']:
                combo = self.data_sources[data_source]['coordinate_combinations'][key]
                coord_display += f" [Combines: {', '.join(combo.get('combines', []))} ({combo.get('operation', 'add')})]"
            listbox.insert(tk.END, coord_display)

    def process_cells_for_automap(self, data_source, cell_addresses, min_score=0.3, destination=None):
        """Unified function to process any collection of cells for automapping"""
        # Get current destination if not provided
        if destination is None:
            destination = self.get_current_destination()
            if destination is None:
                destination = "Default"
        
        print(f"\n=== PROCESS_CELLS_FOR_AUTOMAP DEBUG ===")
        print(f"Input parameters:")
        print(f"  - data_source: {data_source}")
        print(f"  - cell_addresses: {cell_addresses}")
        print(f"  - min_score: {min_score}")
        print(f"  - destination: {destination}")
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
                best_match, score, header = self.auto_map_coordinate_semantic(data_source, target_cell, min_score)
                print(f"  - Semantic mapping result:")
                print(f"    - best_match: {best_match}")
                print(f"    - score: {score}")
                print(f"    - header: {header}")
                
                # Record the result
                if best_match:
                    print(f"  - SUCCESS: Adding mapping for {target_cell}")
                    # Add the mapping - use per-destination structure
                    if destination not in self.data_sources[data_source]['coordinate_values']:
                        self.data_sources[data_source]['coordinate_values'][destination] = {}
                    coordinate_values = self.data_sources[data_source]['coordinate_values'][destination]
                    print(f"  - Current coordinate values count for '{destination}': {len(coordinate_values)}")
                    coordinate_values[target_cell] = best_match
                    self.data_sources[data_source]['coordinate_values'][destination] = coordinate_values
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

