import tkinter as tk
from tkinter import filedialog
from main_functions import *
from tkinter import scrolledtext, Toplevel, Listbox, Button, Frame, MULTIPLE, Checkbutton, IntVar, Label, Scrollbar, Canvas, Entry, Spinbox, END, BOTH, LEFT, RIGHT, TOP, BOTTOM, X, Y, W, E, NW, SE, EW, messagebox, VERTICAL
from tkinter import ttk  # Import ttk for Combobox
import json
import re
import os
import xlwings as xw
from Data_Extraction import DatasheetExtractor
from xlsx_search import ExcelSearchApp
from edit_xlsx import ExcelEditorApp
import pickle
from tkinter.simpledialog import askstring
from excel_manager import *
import openpyxl
from excel_macro_viewer import ExcelMacroViewer # Add this import
from excel_regex_search import ExcelRegexSearchApp


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
        self.root = root
        self.root.title("Datasheet Helper App")
        self.excel_mgr = ExcelManager()
        self.new_sheets = []
        self.td_selected_sheets = None
        self.pc_selected_sheets = None
        self.parameters = {
            'process_conditions_path': '',
            'tag_data_path': '',
            'td_coordinate_values': {},
            'pc_coordinate_values': {},
            'td': {},
            'td_headers': ['TAG NUMBER'],
            'pc': {},
            'pc_headers': ['Line No.'],
            'transformation_code': 'int(x.split("-")[2])',
            'td_xkey': '',
            'tag_filters': [],
            'tag_cell_values': {},
            'datasheet_path': '',
            'source_sheet_name': 'TEMPLATE',
            'datasheet_coord': 'U8',
            'ds_str': 'DS-IA-',
            'rows_per_sheet': 1,
            'top_tag': 'A1',
            'blank_cell_tolerance': 2,
            'sig_figs': 4, # Default significant figures
            'rounding_tolerance': 1e-2, # Default rounding tolerance
            'coordinate_conversions': {}, # Stores coordinate-specific unit conversions
            'coordinate_combinations': {} # Stores coordinate combinations (e.g., {'A1': {'combines': ['B1', 'C1'], 'operation': 'add'}})
        }

        # Initialize all parameters in the __init__ method
        for param, value in self.parameters.items():
            setattr(self, param, value)

        self.create_widgets()

        #self.root.protocol("WM_DELETE_WINDOW", self.on_closing)


    def create_widgets(self):
        # Create menu bar
        self.menu_bar = tk.Menu(self.root)
        self.root.config(menu=self.menu_bar)

        # Create Commands menu
        self.command_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label="Commands", menu=self.command_menu)

        # Add menu items
        menu_commands = [
            ("Load Settings", self.load_settings),
            ("Load Settings except td and pc", self.load_settings_except_td_pc),
            ("Save Settings", self.save_settings),
            ("Populate Index (td) from Json", self.load_td_from_json),
            ("Populate Index (td) from Datasheet", self.load_td_from_datasheet),
            ("Populate Index (pc) from Datasheet", self.load_pc_from_datasheet),
            ("Run xlsx search app", self.open_excel_search_app),
            ("Save and close", self.save_and_close_workbook),
            ("Populate Headers on Datasheets", self.open_edit_xlsx),
            ("View Coordinate Value Data", self.display_coordinate_values),
            ("Delete newly added datasheets", self.delete_added_sheets),
            ("Modify Keys in PC", self.update_pc_keys),
            ("Rebuild tabs", self.refresh_tab_content),
            ("Delete Certain Sheets by Prefix", self.delete_sheets_by_prefix),
            ("Excel Macros", self.open_excel_macros_window), # Add new entry
            ("Excel Regex Search App", self.open_excel_regex_search_app)
        ]

        for label, command in menu_commands:
            self.command_menu.add_command(label=label, command=command)

        # Create main container frame
        main_frame = tk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        self.entries = []  # Store entries for later reference

        # Process Conditions Row
        pc_frame = tk.Frame(main_frame)
        pc_frame.pack(fill=tk.X, pady=5)

        pc_label = tk.Label(pc_frame, text="Process Conditions (Source 2)", width=30)
        pc_label.pack(side=tk.LEFT, padx=5)

        pc_entry = tk.Entry(pc_frame)
        pc_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        self.entries.append((pc_entry, "process_conditions_path"))

        pc_buttons = tk.Frame(pc_frame)
        pc_buttons.pack(side=tk.RIGHT)

        tk.Button(pc_buttons, text="Browse",
                  command=lambda: self.browse(pc_entry, "process_conditions_path")).pack(side=tk.LEFT, padx=2)
        tk.Button(pc_buttons, text="Configure",
                  command=lambda: self.configure("Process Conditions")).pack(side=tk.LEFT, padx=2)
        tk.Button(pc_buttons, text="Generate",
                  command=self.generate_process_conditions).pack(side=tk.LEFT, padx=2)
        tk.Button(pc_buttons, text="View",
                  command=lambda: self.view_data("Process Conditions")).pack(side=tk.LEFT, padx=2)

        # Instrument Index Row
        ii_frame = tk.Frame(main_frame)
        ii_frame.pack(fill=tk.X, pady=5)

        ii_label = tk.Label(ii_frame, text="Instrument Index (Source 1)", width=30)
        ii_label.pack(side=tk.LEFT, padx=5)

        ii_entry = tk.Entry(ii_frame)
        ii_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        self.entries.append((ii_entry, "tag_data_path"))

        ii_buttons = tk.Frame(ii_frame)
        ii_buttons.pack(side=tk.RIGHT)

        tk.Button(ii_buttons, text="Browse",
                  command=lambda: self.browse(ii_entry, "tag_data_path")).pack(side=tk.LEFT, padx=2)
        tk.Button(ii_buttons, text="Configure",
                  command=lambda: self.configure("Instrument Index")).pack(side=tk.LEFT, padx=2)
        tk.Button(ii_buttons, text="Generate",
                  command=self.generate_tag_data).pack(side=tk.LEFT, padx=2)
        tk.Button(ii_buttons, text="View",
                  command=lambda: self.view_data("Instrument Index")).pack(side=tk.LEFT, padx=2)

        # Datasheets Row
        ds_frame = tk.Frame(main_frame)
        ds_frame.pack(fill=tk.X, pady=5)

        ds_label = tk.Label(ds_frame, text="Datasheets (Destination)", width=30)
        ds_label.pack(side=tk.LEFT, padx=5)

        ds_entry = tk.Entry(ds_frame)
        ds_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        self.entries.append((ds_entry, "datasheets"))

        ds_buttons = tk.Frame(ds_frame)
        ds_buttons.pack(side=tk.RIGHT)

        tk.Button(ds_buttons, text="Browse",
                  command=lambda: self.browse(ds_entry, "datasheets")).pack(side=tk.LEFT, padx=2)
        tk.Button(ds_buttons, text="Configure",
                  command=lambda: self.configure("Datasheets")).pack(side=tk.LEFT, padx=2)
        tk.Button(ds_buttons, text="Generate",
                  command=self.add_datasheets).pack(side=tk.LEFT, padx=2)
        tk.Button(ds_buttons, text="View",
                  command=lambda: self.view_data("Datasheets")).pack(side=tk.LEFT, padx=2)

        # Create notebook after entries
        self.config_frame = tk.Frame(main_frame)
        self.config_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        self.notebook = ttk.Notebook(self.config_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        # Create tab pages
        coordinates_tab = ttk.Frame(self.notebook)
        datasheet_tab = ttk.Frame(self.notebook)
        filters_tab = ttk.Frame(self.notebook)
        transform_tab = ttk.Frame(self.notebook)

        self.notebook.add(coordinates_tab, text='Coordinates')
        self.notebook.add(datasheet_tab, text='Datasheet')
        self.notebook.add(filters_tab, text='Filters')
        self.notebook.add(transform_tab, text='Transform')

        self.create_coordinates_tab(coordinates_tab)
        self.create_datasheet_tab(datasheet_tab)
        self.create_filters_tab(filters_tab)
        self.create_transform_tab(transform_tab)

    # endregion
    # region Tab Creation

    def init_excel(self):
        """Initialize Excel only when needed"""
        if self.datasheet_path:
            try:
                # Check if workbook reference is still valid
                if self.excel_mgr.wb and self.excel_mgr.wb.name:
                    return
            except:
                # Workbook was closed, reset references
                self.excel_mgr.wb = None
                self.excel_mgr.app = None
            
            self.excel_mgr.open_workbook(self.datasheet_path)

    def create_coordinates_tab(self, tab):

        def get_combo_values():
            td_values = []
            pc_values = []
            for key, value in self.td.items():
                td_values = list(value.keys())
                break
            for key, value in self.pc.items():
                pc_values = list(value.keys())
                break
            return td_values, pc_values

        def add_coordinate(coord_type, entry, combo, listbox):
            coord = entry.get()
            value = combo.get()
            if coord and value:
                if coord_type == "td":
                    self.td_coordinate_values[coord] = value
                    if not self.td_coordinate_values:
                        self.top_tag = coord
                else:
                    self.pc_coordinate_values[coord] = value
                update_listboxes()

        def remove_coordinate(coord_type, listbox):
            selected = listbox.curselection()
            if selected:
                idx = selected[0]
                coord = listbox.get(idx).split(':')[0]
                if coord_type == "td":
                    del self.td_coordinate_values[coord]
                else:
                    del self.pc_coordinate_values[coord]
                update_listboxes()

        def clear_coordinates(coord_type, listbox):
            if coord_type == "td":
                self.td_coordinate_values.clear()
            else:
                self.pc_coordinate_values.clear()
            update_listboxes()

        def update_entry():
            if self.datasheet_path:
                try:
                    current_selection = xw.apps.active.selection.address
                    current_selection = current_selection.split(':')[0].replace('$', '')
                    entry_var.set(current_selection)
                except:
                    pass
            tab.after(200, update_entry)

        def update_td_listbox():
            td_listbox.delete(0, tk.END)
            for key, value in self.td_coordinate_values.items():
                # Add placeholder for conversion details
                coord_display = f"{key}: {value}"
                if key in self.coordinate_conversions:
                    conv = self.coordinate_conversions[key]
                    coord_display += f" [{conv.get('in_unit', '?')}->{conv.get('out_unit', '?')}]"
                if key in self.coordinate_combinations:
                    combo = self.coordinate_combinations[key]
                    coord_display += f" [Combines: {', '.join(combo.get('combines', []))} ({combo.get('operation', 'add')})]"
                td_listbox.insert(tk.END, coord_display)
            try:
                first_entry = td_listbox.get(0)
                self.top_tag = first_entry.split(':')[0]
            except:
                print('failed to set top_tag')

        def update_pc_listbox():
            pc_listbox.delete(0, tk.END)
            for key, value in self.pc_coordinate_values.items():
                # Add placeholder for conversion details
                coord_display = f"{key}: {value}"
                if key in self.coordinate_conversions:
                    conv = self.coordinate_conversions[key]
                    coord_display += f" [{conv.get('in_unit', '?')}->{conv.get('out_unit', '?')}]"
                if key in self.coordinate_combinations:
                    combo = self.coordinate_combinations[key]
                    coord_display += f" [Combines: {', '.join(combo.get('combines', []))} ({combo.get('operation', 'add')})]"
                pc_listbox.insert(tk.END, coord_display)

        def update_listboxes():
            update_td_listbox()
            update_pc_listbox()

        def reinitialize():
            #init_excel()
            td_values, pc_values = get_combo_values()
            td_combo['values'] = td_values
            pc_combo['values'] = pc_values
            update_listboxes()

        # Initial Excel setup if path exists
        #init_excel()
        td_combo_values, pc_combo_values = get_combo_values()

        # UI Setup (same as before)
        top_frame = ttk.Frame(tab)
        top_frame.pack(fill="x", padx=10, pady=5)

        coord_label = ttk.Label(top_frame, text="Enter Key Coordinate:\n(First entry is top tag default)")
        coord_label.pack(side="left")

        entry_var = tk.StringVar()
        coord_entry = ttk.Entry(top_frame, textvariable=entry_var)
        coord_entry.pack(side="left", fill="x", expand=True)

        content_frame = ttk.Frame(tab)
        content_frame.pack(fill="both", expand=True, padx=10, pady=5)

        td_frame = ttk.LabelFrame(content_frame, text="TD Coordinates")
        td_frame.pack(side="left", fill="both", expand=True, padx=5)

        td_controls = ttk.Frame(td_frame)
        td_controls.pack(fill="x", padx=5, pady=5)

        td_label = ttk.Label(td_controls, text="Select TD Value:")
        td_label.pack(side="left")

        td_combo = ttk.Combobox(td_controls, values=td_combo_values, state="readonly")
        td_combo.pack(side="left", fill="x", expand=True, padx=5)

        td_btn_frame = ttk.Frame(td_frame)
        td_btn_frame.pack(fill="x", padx=5)

        td_listbox = tk.Listbox(td_frame, height=15)
        td_listbox.pack(fill="both", expand=True, padx=5, pady=5)

        ttk.Button(td_btn_frame, text="Add to TD",
                   command=lambda: add_coordinate("td", coord_entry, td_combo, td_listbox)).pack(side="left", padx=2)
        ttk.Button(td_btn_frame, text="Remove",
                   command=lambda: remove_coordinate("td", td_listbox)).pack(side="left", padx=2)
        ttk.Button(td_btn_frame, text="Clear All",
                   command=lambda: clear_coordinates("td", td_listbox)).pack(side="left", padx=2)

        xfn_button = ttk.Button(content_frame, text="Transformation\nCode and Key", command=lambda: self.notebook.select(3))
        xfn_button.pack(side="left", padx=10)

        pc_frame = ttk.LabelFrame(content_frame, text="PC Coordinates")
        pc_frame.pack(side="left", fill="both", expand=True, padx=5)

        pc_controls = ttk.Frame(pc_frame)
        pc_controls.pack(fill="x", padx=5, pady=5)

        pc_label = ttk.Label(pc_controls, text="Select PC Value:")
        pc_label.pack(side="left")

        pc_combo = ttk.Combobox(pc_controls, values=pc_combo_values, state="readonly")
        pc_combo.pack(side="left", fill="x", expand=True, padx=5)

        pc_btn_frame = ttk.Frame(pc_frame)
        pc_btn_frame.pack(fill="x", padx=5)

        pc_listbox = tk.Listbox(pc_frame, height=15)
        pc_listbox.pack(fill="both", expand=True, padx=5, pady=5)

        # --- Context Menu Setup ---
        self.coord_context_menu = tk.Menu(tab, tearoff=0)
        self.coord_context_menu.add_command(label="Change Source Key", command=self.open_change_source_key_dialog)
        self.coord_context_menu.add_separator()
        self.coord_context_menu.add_command(label="Add/Edit Conversion", command=self.open_conversion_dialog)
        self.coord_context_menu.add_command(label="Remove Conversion", command=self.remove_conversion)
        self.coord_context_menu.add_command(label="Combine", command=self.open_combination_dialog)
        self.coord_context_menu.add_command(label="Remove Combination", command=self.remove_combination)
        self.coord_context_menu.add_separator()
        self.coord_context_menu.add_command(label="Cancel")

        self.selected_coord_for_context = None # Store the coordinate clicked on

        def show_coord_context_menu(event, listbox_widget, coord_type):
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

        # Bind right-click to both listboxes
        td_listbox.bind("<Button-3>", lambda event: show_coord_context_menu(event, td_listbox, 'td'))
        pc_listbox.bind("<Button-3>", lambda event: show_coord_context_menu(event, pc_listbox, 'pc'))
        # ---

        ttk.Button(pc_btn_frame, text="Add to PC",
                   command=lambda: add_coordinate("pc", coord_entry, pc_combo, pc_listbox)).pack(side="left", padx=2)
        ttk.Button(pc_btn_frame, text="Remove",
                   command=lambda: remove_coordinate("pc", pc_listbox)).pack(side="left", padx=2)
        ttk.Button(pc_btn_frame, text="Clear All",
                   command=lambda: clear_coordinates("pc", pc_listbox)).pack(side="left", padx=2)

        update_listboxes()
        tab.after(200, update_entry)

        # Expose reinitialize method for external calls
        tab.reinitialize = reinitialize
        return tab

    def create_datasheet_tab(self, tab):
        fields = [
            ("Source Sheet Name (Leave blank to only update existing)", "source_sheet_name", ttk.Combobox),
            ("Datasheet Coordinate", "datasheet_coord", ttk.Entry),
            ("Datasheet Prefix used to identify sheets that should be updated and fill DS numbers\n(If blank, we will look for existing tags in all sheets starting from the top tag)\nThis could be problematic if there is a cover sheet we try pulling tags from which might cause issues with accessing sheets that dont exist", "ds_str", ttk.Entry),
            #("Tag Pattern", "tag_pattern", ttk.Entry),
            ("Top Tag", "top_tag", ttk.Entry),
            ("Rows per Sheet", "rows_per_sheet", ttk.Entry)
        ]

        entries = {}
        for i, (label, attr, widget_type) in enumerate(fields):
            ttk.Label(tab, text=label).grid(row=i, column=0, padx=5, pady=5, sticky="w")
            w = widget_type(tab)
            w.grid(row=i, column=1, padx=5, pady=5, sticky="ew")
            if hasattr(self, attr):
                w.insert(0, getattr(self, attr))
            if widget_type == ttk.Combobox:
                w['values'] = self.get_sheet_names()
            entries[attr] = w

        # Add Sig Figs and Tolerance fields
        i = len(fields) # Get the next row index
        ttk.Label(tab, text="Significant Figures for Rounding").grid(row=i, column=0, padx=5, pady=5, sticky="w")
        sig_figs_entry = ttk.Entry(tab)
        sig_figs_entry.grid(row=i, column=1, padx=5, pady=5, sticky="ew")
        sig_figs_entry.insert(0, getattr(self, 'sig_figs'))
        entries['sig_figs'] = sig_figs_entry

        i += 1
        ttk.Label(tab, text="Rounding Tolerance (e.g., 1e-2)").grid(row=i, column=0, padx=5, pady=5, sticky="w")
        tolerance_entry = ttk.Entry(tab)
        tolerance_entry.grid(row=i, column=1, padx=5, pady=5, sticky="ew")
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
            print("Datasheet settings saved.") # Add confirmation

        save_btn = ttk.Button(tab, text="Save", command=save_datasheet_settings)
        save_btn.grid(row=len(fields) + 2, column=0, columnspan=2, pady=10) # Adjust row index for the button
        tab.grid_columnconfigure(1, weight=1)

    def create_filters_tab(self, tab):
        # Add filter frame

        # Add explanatory label for filter functionality
        filter_info_label = ttk.Label(tab, text="Can use comma to include multiple filter terms")
        filter_info_label.pack(anchor="w", padx=5, pady=(5, 0))
        add_frame = ttk.Frame(tab)
        add_frame.pack(fill="x", padx=5, pady=5)

        filters_entries = []
        td_combo_values = []

        for key, value in self.td.items():
            td_combo_values = list(value.keys())
            break

        def add_filter_row(name='', filter_value=''):
            new_row = len(filters_entries) + 1
            name_label = tk.Label(content_frame, text=f"Index Key {new_row}:")
            name_label.grid(row=new_row, column=0)
            name_entry = ttk.Combobox(content_frame, values=td_combo_values)
            name_entry.grid(row=new_row, column=1)
            name_entry.set(name)

            filter_label = tk.Label(content_frame, text=f"Filter {new_row}:")
            filter_label.grid(row=new_row, column=2)
            filter_entry = tk.Entry(content_frame)
            filter_entry.grid(row=new_row, column=3)
            filter_entry.insert(0, filter_value)

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
        # Source key selector
        key_frame = ttk.Frame(tab)
        key_frame.pack(fill="x", padx=5, pady=5)

        ttk.Label(key_frame, text="Index Source Key:").pack(side="left")

        td_combo_values = []
        for key, value in self.td.items():
            td_combo_values = list(value.keys())
            break

        key_combo = ttk.Combobox(key_frame, values=td_combo_values)
        key_combo.pack(side="left", fill="x", expand=True, padx=5)
        if self.td_xkey:
            key_combo.set(self.td_xkey)

        # Transformation code entry
        code_frame = ttk.Frame(tab)
        code_frame.pack(fill="x", padx=5, pady=5)

        ttk.Label(code_frame, text="PC key, x = Index Value:").pack(side="left")
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
            self.td_xkey = key_combo.get()
            self.transformation_code = code_entry.get()

        ttk.Button(tab, text="Save", command=save_transform).pack(pady=5)

    def refresh_tab_content(self):
        # Recreate all tabs with updated data
        for tab in self.notebook.winfo_children():
            tab.destroy()

        coordinates_tab = ttk.Frame(self.notebook)
        datasheet_tab = ttk.Frame(self.notebook)
        filters_tab = ttk.Frame(self.notebook)
        transform_tab = ttk.Frame(self.notebook)

        self.notebook.add(coordinates_tab, text='Coordinates')
        self.notebook.add(datasheet_tab, text='Datasheet')
        self.notebook.add(filters_tab, text='Filters')
        self.notebook.add(transform_tab, text='Transform')

        self.create_coordinates_tab(coordinates_tab)
        self.create_datasheet_tab(datasheet_tab)
        self.create_filters_tab(filters_tab)
        self.create_transform_tab(transform_tab)

    # endregion

    # region Configuration Windows

    def get_xkey(self):
        xkey_window = tk.Toplevel(self.root)
        xkey_window.title("Configure translation_lambda")

        frame1 = tk.Frame(xkey_window)
        frame1.pack(fill="x", pady=5, padx=5)
        frame2 = tk.Frame(xkey_window, pady=5, padx=5)
        frame2.pack(fill="x")

        # Label for "Index Source Key"
        td_label = tk.Label(frame1, text="Index Source Key: ")
        td_label.pack(side='left', padx=5)

        # Dropdown options for ttk combobox
        td_combo_values = []
        # Example values for the combo box
        for key, value in self.td.items():
            td_combo_values = list(value.keys())
            print(list(td_combo_values))
            break

        # Creating the ttk combobox
        td_combo = ttk.Combobox(frame1, values=td_combo_values)
        td_combo.insert(0, self.td_xkey)
        td_combo.pack(side='left', fill="x", expand=True)

        # Label for "'x' is the source. Code: "
        code_label = tk.Label(frame2, text="PC key, x = Index Value: ")
        code_label.pack(side='left', padx=5)

        # Entry box for transformation code
        transformation_entry = tk.Entry(frame2)
        transformation_entry.insert(0, self.transformation_code)
        transformation_entry.pack(side='left', fill="x", expand=True)

        # Save button
        def save_values():
            # Get the selected value from the Combobox
            selected_td_xkey = td_combo.get()
            # Update self.td_xkey with the selected value
            self.td_xkey = selected_td_xkey

            # Get the text entered in the Entry widget
            entered_transformation_code = transformation_entry.get()
            # Update self.transformation_code with the entered text
            self.transformation_code = entered_transformation_code

            # Close the Toplevel window
            xkey_window.destroy()

        save_button = tk.Button(xkey_window, text="Save", command=save_values)
        save_button.pack(expand=True)

        # Run the tkinter main loop
        xkey_window.mainloop()

    def set_tag_filters(self):
        view_window = tk.Toplevel(self.root)
        view_window.title("Set Tag Filters (Comma for OR). ReGenerate Coordinates if necessary")

        filters_entries = []

        td_combo_values = []
        # Example values for the combo box
        for key, value in self.td.items():
            td_combo_values = list(value.keys())
            print(list(td_combo_values))
            break

        def add_filter_row(name='', filter_value=''):
            new_row = len(filters_entries) + 1

            name_label = tk.Label(view_window, text=f"Index Key {new_row}:")
            name_label.grid(row=new_row, column=0)
            name_entry = ttk.Combobox(view_window, values=td_combo_values)
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

    def generate_process_conditions(self):
        # Pass the tolerance value
        self.pc = generate_dictionary_from_xlsx(self.process_conditions_path, self.pc_headers,
                                              parent=self.root, selected_sheets=self.pc_selected_sheets,
                                              max_empty_allowed=self.blank_cell_tolerance)
        if self.pc is not None: # Check if generation was successful (not cancelled)
             show_nested_dict_analysis(self.pc)
             print("Generated Process Conditions")
             self.refresh_tab_content()
        else:
             print("Process Conditions generation cancelled or failed.")

    def generate_tag_data(self):
        print("Generating Tag Data")
        print(self.tag_data_path)
        print(self.td_headers)
        # Pass the tolerance value
        self.td = generate_dictionary_from_xlsx(self.tag_data_path, self.td_headers,
                                              parent=self.root, selected_sheets=self.td_selected_sheets,
                                              max_empty_allowed=self.blank_cell_tolerance)
        if self.td is not None: # Check if generation was successful
             show_nested_dict_analysis(self.td)
             print("Generated Tag Data")
             self.refresh_tab_content()
        else:
             print("Tag Data generation cancelled or failed.")

    def assign_value_coordinate_to_tag(self):
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
        
        for tag in self.td:
            if tag:
                # filter out
                continue_flag = False
                for header, filter_key in self.tag_filters:
                    print('tag', tag)
                    print('header', header)
                    # Split the filter key on commas to get multiple acceptable values
                    acceptable_values = [value.strip() for value in filter_key.split(',')]

                    # If the tag's value for this header isn't in our acceptable values, filter it out
                    if self.td[tag][header] not in acceptable_values:
                        continue_flag = True
                        break

                if continue_flag:
                    continue

                # Process TD coordinates
                data = {}
                for coordinate, value in self.td_coordinate_values.items():
                    print(f'tag {tag}, value {value}, coord {coordinate}')
                    raw_value = self.td[tag].get(value) # Use .get() for safety
                    processed_value = process_coordinate_value(coordinate, raw_value)
                    data[coordinate] = processed_value
                    if processed_value is None:
                        print(f"  Warning: Key '{value}' not found in td for tag '{tag}'. Skipping coordinate '{coordinate}'.")

                # Apply combinations to TD data
                apply_combinations(data)
                
                try:
                    interface = translate(self.td[tag][self.td_xkey], self.transformation_code)
                    print(f'tag: {tag}, td_xkey: {self.td_xkey}, interface: {interface}')
                    print("length ", len(self.pc_coordinate_values))

                    for coordinate, value in self.pc_coordinate_values.items():
                        print("value ", value)
                        try:
                            raw_value = self.pc[interface].get(value) # Use .get() for safety
                            processed_value = process_coordinate_value(coordinate, raw_value)
                            data[coordinate] = processed_value
                            if processed_value is None:
                                print(f"  Warning: Key '{value}' not found in pc for interface '{interface}'. Skipping coordinate '{coordinate}'.")
                        except KeyError:
                             print(f"  Warning: Interface key '{interface}' not found in pc. Skipping PC coordinate '{coordinate}'.")
                             data[coordinate] = None # Or handle as needed
                        except Exception as e:
                            print(f'Error processing PC coordinate {coordinate} for interface {interface}: {e}')
                            data[coordinate] = None # Or handle as needed

                except KeyError as e:
                    print(f'xkey pc interface fail: Key \'{e}\' not found for tag \'{tag}\'.')
                except Exception as e:
                    print(f'xkey pc interface fail for tag {tag}: {e}')

                # Apply combinations to PC data as well
                apply_combinations(data)

                self.tag_cell_values[tag] = data

        print("Coordinate Values generated:", self.tag_cell_values)

    def add_datasheets(self):
        print('assigning tag coordinates')
        self.assign_value_coordinate_to_tag()
        print("Adding/Updating Datasheets")
        
        try:
            # Force reinitialization of Excel connection
            self.init_excel()
            print('Excel initialized')
            # Additional check for valid workbook reference
            if not self.excel_mgr.wb:
                raise Exception("Excel workbook not properly initialized")

            # lets print all the variables that go into add_update_datasheets
            print('source_sheet_name', self.source_sheet_name)
            print('tag_cell_values', self.tag_cell_values)
            print('datasheet_coord', self.datasheet_coord)
            print('ds_str', self.ds_str)
            print('rows_per_sheet', self.rows_per_sheet)
            print('top_tag', self.top_tag)

            # Check for potential naming conflict
            if self.source_sheet_name and self.source_sheet_name.startswith(self.ds_str):
                print(f"WARNING: Source sheet name '{self.source_sheet_name}' starts with datasheet prefix '{self.ds_str}'")
                print("This may cause issues with sheet identification. Consider using a different prefix.")

            self.new_sheets = add_update_datasheets(self.excel_mgr.wb, self.source_sheet_name,
                                            self.tag_cell_values, self.datasheet_coord,
                                            self.ds_str, rows_per_sheet=self.rows_per_sheet,
                                            key_coordinate=self.top_tag,
                                            sig_figs=self.sig_figs, # Pass sig_figs
                                            tolerance=self.rounding_tolerance) # Pass tolerance
            self.excel_mgr.mark_as_modified()
            print("DONE")
        except Exception as e:
            print(f"Excel connection error: {e}")
            messagebox.showerror("Excel Connection Error", 
                               "Excel connection lost. Please ensure Excel is open and try again.")
            # Reset Excel connection
            self.excel_mgr.wb = None
            self.excel_mgr.app = None

    # endregion

    # region Data Loading and Saving

    def load_settings(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Settings files", "*.json *.pkl"), ("JSON files", "*.json"), ("Pickle files", "*.pkl")]
        )

        if not file_path:
            return

        try:
            file_ext = os.path.splitext(file_path)[1].lower()
            with open(file_path, 'rb' if file_ext == '.pkl' else 'r') as file:
                settings_data = pickle.load(file) if file_ext == '.pkl' else json.load(file)

                for key, value in settings_data.items():
                    if key in self.parameters:
                        setattr(self, key, value)

                # Force tab content refresh - add parentheses to actually call the method
                self.refresh_tab_content()
                self.update_entries()
                print(f"Settings loaded successfully from {file_ext} file! You may need to restart Excel")

        except Exception as e:
            print(f"Error loading settings: {e}")

    def load_settings_except_td_pc(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Settings files", "*.json *.pkl"), ("JSON files", "*.json"), ("Pickle files", "*.pkl")]
        )

        if not file_path:
            return

        # Define excluded parameters
        excluded_params = {'td', 'pc'}

        try:
            file_ext = os.path.splitext(file_path)[1].lower()
            with open(file_path, 'rb' if file_ext == '.pkl' else 'r') as file:
                settings_data = pickle.load(file) if file_ext == '.pkl' else json.load(file)

                for key, value in settings_data.items():
                    if key in self.parameters and key not in excluded_params:
                        setattr(self, key, value)

                # Force tab content refresh - add parentheses to actually call the method
                self.refresh_tab_content()
                self.update_entries()
                print(f"Settings loaded successfully from {file_ext} file!")

        except Exception as e:
            print(f"Error loading settings: {e}")

    def save_settings(self, use_pickle=True):
        """
        Saves settings to either JSON or pickle file.
        Args:
            use_pickle (bool): If True, saves as pickle file. Use only if JSON serialization fails.
        """

        file_ext = ".pkl" if use_pickle else ".json"
        file_path = filedialog.asksaveasfilename(
            defaultextension=file_ext,
            filetypes=[("Settings files", f"*{file_ext}")]
        )

        if not file_path:
            return

        settings_to_save = {}
        for key in self.parameters:
            value = getattr(self, key)
            settings_to_save[key] = value

        try:
            if use_pickle:
                with open(file_path, 'wb') as file:
                    pickle.dump(settings_to_save, file)
            else:
                with open(file_path, 'w') as file:
                    json.dump(settings_to_save, file, indent=4)

            print(f"Settings saved successfully as {file_ext}!")

        except TypeError as e:
            if not use_pickle:
                print(f"JSON serialization failed: {e}")
                print("Try saving as pickle file instead.")
            else:
                print(f"Pickle serialization failed: {e}")
        except Exception as e:
            print(f"Error occurred while saving settings: {e}")

    def load_td_from_datasheet(self):
        def set_td(td):
            self.td = td

        app_window = tk.Toplevel(root)
        DatasheetExtractor(app_window, callback=set_td)

    def load_pc_from_datasheet(self):
        def set_pc(pc):
            self.pc = pc

        app_window = tk.Toplevel(root)
        DatasheetExtractor(app_window, callback=set_pc)

    def load_td_from_json(self):
        # use tkinter to ask for the json file path
        # use load_dict_from_json(file_path) to set self.td
        # Ask the user to select a JSON file
        file_path = filedialog.askopenfilename(title="Select JSON file", filetypes=[("JSON files", "*.json")])

        # Check if a file was selected
        if file_path:
            # Load the JSON file using load_dict_from_json
            self.td = load_dict_from_json(file_path)

    # endregion

    # region Data Display

    def view_data(self, text):

        print(f"Viewing {text}")

        if text == "Process Conditions":
            # open a new tkinter popup window with a scroll bar showing all the key-value pairs in the self.pc dictionary
            self.display_process_conditions()
        if text == "Instrument Index":
            # open a new tkinter popup window with a scroll bar showing all the key-value pairs in the self.pc dictionary
            self.display_tag_data()
        if text == "Coordinate-Value Data":
            # open a new tkinter popup window with a scroll bar showing all the key-value pairs in the self.pc dictionary
            self.display_coordinate_values()
        if text == "Datasheets":
            # open a new tkinter popup window with a scroll bar showing all the key-value pairs in the self.pc dictionary
            os.startfile(self.datasheet_path)

    def display_process_conditions(self):
        # Create a new window
        view_window = tk.Toplevel(self.root)
        view_window.title("Process Conditions")

        # Create a scrolled text widget to display the process conditions
        scrolled_text = scrolledtext.ScrolledText(view_window, width=40, height=20)
        scrolled_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)  # Fill and expand to fill the window

        # Display the process conditions content in the scrolled text widget
        if self.pc:
            for key, value in self.pc.items():
                scrolled_text.insert(tk.END, f"{key}: {value}\n\n")
        else:
            scrolled_text.insert(tk.END, "No process conditions data available.")

        scrolled_text.configure(state='disabled')  # Make

    def display_tag_data(self):
        # Create a new window
        view_window = tk.Toplevel(self.root)
        view_window.title("Instrument Index Tag Data")

        # Create a scrolled text widget to display the process conditions
        scrolled_text = scrolledtext.ScrolledText(view_window, width=40, height=20)
        scrolled_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)  # Fill and expand to fill the window

        # Display the process conditions content in the scrolled text widget
        if self.td:
            for key, value in self.td.items():
                scrolled_text.insert(tk.END, f"{key}: {value}\n\n")
        else:
            scrolled_text.insert(tk.END, "No tag data available.")

        scrolled_text.configure(state='disabled')  # Make

    def display_coordinate_values(self):

        # Create a new window
        view_window = tk.Toplevel(self.root)
        view_window.title("Coordinate values")

        td_label = tk.Label(view_window, text='Index Tag data Coordinates')
        scrolled_text1 = scrolledtext.ScrolledText(view_window, width=40, height=20)
        td_label.pack()
        scrolled_text1.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)  # Fill and expand to fill the window

        # Display the process conditions content in the scrolled text widget
        if self.tag_cell_values:
            for key, value in self.tag_cell_values.items():
                scrolled_text1.insert(tk.END, f"{key}: {value}\n\n")
        else:
            scrolled_text1.insert(tk.END, "No tag data available.")

        scrolled_text1.configure(state='disabled')  # Make

    def display_coordinate_values(self):

        # Create a new window
        view_window = tk.Toplevel(self.root)
        view_window.title("Coordinate values")

        td_label = tk.Label(view_window, text='Index Tag data Coordinates')
        scrolled_text1 = scrolledtext.ScrolledText(view_window, width=40, height=20)
        td_label.pack()
        scrolled_text1.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)  # Fill and expand to fill the window

        # Display the process conditions content in the scrolled text widget
        if self.tag_cell_values:
            for key, value in self.tag_cell_values.items():
                scrolled_text1.insert(tk.END, f"{key}: {value}\n\n")
        else:
            scrolled_text1.insert(tk.END, "No tag data available.")

        scrolled_text1.configure(state='disabled')  # Make

    # endregion

    # region Excel Operations

    def save_and_close_workbook(self):
        """Save and close the current workbook"""
        self.excel_mgr.close_workbook()

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

    def get_sheet_names(self):
        if not self.datasheet_path:
            return []
        try:
            wb = openpyxl.load_workbook(self.datasheet_path, read_only=True)
            return wb.sheetnames
        except Exception as e:
            print(f"Error getting sheet names: {e}")
            return []

    def on_closing(self):
        """Handle application closing"""
        if self.excel_mgr.is_dirty:
            if messagebox.askyesno("Save Changes",
                                 "There are unsaved changes. Would you like to save before closing?"):
                self.excel_mgr.save_workbook()
        self.excel_mgr.cleanup()
        self.root.destroy()


    # endregion

    # region UI Utilities

    def update_entry(self, entry, variable):
        filename = ''
        if variable == "process_conditions_path":
            filename = self.process_conditions_path
        elif variable == "tag_data_path":
            filename = self.tag_data_path
        elif variable == "coordinate_value":
            filename = self.coordinate_value_path
        elif variable == "datasheets":
            filename = self.datasheet_path

        entry.delete(0, tk.END)
        entry.insert(0, filename)

    def update_entries(self):
        for entry, entry_var in self.entries:
            try:
                self.update_entry(entry, entry_var)
                entry.xview_moveto(1)
            except Exception as e:
                print(f'error {e}')

    def browse(self, entry, variable):
        filename = filedialog.askopenfilename()
        entry.delete(0, tk.END)
        entry.insert(0, filename)

        if variable == "process_conditions_path":
            self.process_conditions_path = filename
        elif variable == "tag_data_path":
            self.tag_data_path = filename
        elif variable == "coordinate_value":
            self.coordinate_value_path = filename
        elif variable == "datasheets":
            self.datasheet_path = filename
            print(self.datasheet_path)

    def configure(self, text):
        print(f"Configure {text}")

        if text == 'Instrument Index':
            file_path = self.tag_data_path
            current_headers = self.td_headers
            current_selection = self.td_selected_sheets
            current_tolerance = self.blank_cell_tolerance

            if not file_path or not os.path.exists(file_path):
                messagebox.showwarning("File Not Found", "Please select a valid Instrument Index file first.", parent=self.root)
                return

            result = configure_source_data_dialog(self.root, "Configure Instrument Index Source",
                                                  current_headers, file_path, current_selection, current_tolerance)

            if result:
                self.td_headers = result["headers"]
                self.td_selected_sheets = result["selected_sheets"]
                self.blank_cell_tolerance = result["tolerance"]
                print("Instrument Index configuration updated.")
                # Optionally regenerate data immediately or prompt user
                # self.generate_tag_data() # Example: Regenerate

        elif text == 'Process Conditions':
            file_path = self.process_conditions_path
            current_headers = self.pc_headers
            current_selection = self.pc_selected_sheets
            current_tolerance = self.blank_cell_tolerance # Use the same tolerance setting for now

            if not file_path or not os.path.exists(file_path):
                messagebox.showwarning("File Not Found", "Please select a valid Process Conditions file first.", parent=self.root)
                return

            # Reuse the same dialog function
            result = configure_source_data_dialog(self.root, "Configure Process Conditions Source",
                                                  current_headers, file_path, current_selection, current_tolerance)

            if result:
                self.pc_headers = result["headers"]
                self.pc_selected_sheets = result["selected_sheets"]
                self.blank_cell_tolerance = result["tolerance"] # Update tolerance based on this config too
                print("Process Conditions configuration updated.")
                # Optionally regenerate data immediately or prompt user
                # self.generate_process_conditions() # Example: Regenerate

        elif text == 'Datasheets':
             self.configure_ds() # Keep existing Datasheet config separate
        else:
             print(f"Unknown configuration type: {text}")

    def update_pc_keys(self):
        code = askstring("Enter transformation code for keys in dictionary", "Enter transformation code",
                         initialvalue='"-".join(x.split("-")[-2:])')
        self.pc = transform_dictionary(self.pc, code)

    def configure_ds(self):
        self.init_excel()
        self.refresh_tab_content()
        self.update_entries()

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

    def configure_ds(self):
        self.init_excel()
        self.refresh_tab_content()
        self.update_entries()

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
        all_coordinates = list(self.td_coordinate_values.keys()) + list(self.pc_coordinate_values.keys())
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

    def open_change_source_key_dialog(self):
        if not self.selected_coord_for_context:
            return

        coord = self.selected_coord_for_context
        
        # Determine which coordinate type this is (td or pc)
        coord_type = None
        current_source_key = None
        available_keys = []
        
        if coord in self.td_coordinate_values:
            coord_type = "td"
            current_source_key = self.td_coordinate_values[coord]
            # Get available TD keys
            for key, value in self.td.items():
                available_keys = list(value.keys())
                break
        elif coord in self.pc_coordinate_values:
            coord_type = "pc"
            current_source_key = self.pc_coordinate_values[coord]
            # Get available PC keys
            for key, value in self.pc.items():
                available_keys = list(value.keys())
                break
        else:
            print(f"Coordinate {coord} not found in either TD or PC coordinate values")
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
                if coord_type == "td":
                    self.td_coordinate_values[coord] = new_key
                else:  # pc
                    self.pc_coordinate_values[coord] = new_key
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


if __name__ == "__main__":
    root = tk.Tk()
    app = DatasheetGeneratorApp(root)
    root.mainloop()
