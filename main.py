import tkinter as tk
from tkinter import filedialog
from main_functions import *
from tkinter import scrolledtext
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


class DatasheetGeneratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Data Generator App")
        self.wb = None
        self.app = None
        self.new_sheets = []
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
            'tag_pattern': r'^[0-9]{3}-[A-Z]{2,3}-[0-9]{4}[A-Z]?$',
            'rows_per_sheet': 1,
            'top_tag': 'A1'
        }

        # Initialize all parameters in the __init__ method
        for param, value in self.parameters.items():
            setattr(self, param, value)

        self.create_widgets()

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
            ("Transformation Code and Key", self.get_xkey),
            ("Index Filters", self.set_tag_filters),
            ("Populate Index (td) from Json", self.load_td_from_json),
            ("Populate Index (td) from Datasheet", self.load_td_from_datasheet),
            ("Populate Index (pc) from Datasheet", self.load_pc_from_datasheet),
            ("Update Datasheet", self.update_datasheet),
            ("Run xlsx search app", self.open_excel_search_app),
            ("Save and close", self.save_and_close_workbook),
            ("Populate Headers on Datasheets", self.open_edit_xlsx),
            ("Assign Coordinate Values", self.assign_value_coordinate_to_tag),
            ("View Coordinate Value Data", self.display_coordinate_values),
            ("Configure Datasheet", self.configure_add_datasheet),
            ("Delete newly added datasheets", self.delete_added_sheets),
            ("Modify Keys in PC", self.update_pc_keys)
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

        pc_label = tk.Label(pc_frame, text="Process Conditions", width=15)
        pc_label.pack(side=tk.LEFT, padx=5)

        pc_entry = tk.Entry(pc_frame)
        pc_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        self.entries.append((pc_entry, "process_conditions"))

        pc_buttons = tk.Frame(pc_frame)
        pc_buttons.pack(side=tk.RIGHT)

        tk.Button(pc_buttons, text="Browse",
                  command=lambda: self.browse(pc_entry, "process_conditions")).pack(side=tk.LEFT, padx=2)
        tk.Button(pc_buttons, text="Configure",
                  command=lambda: self.configure("Process Conditions")).pack(side=tk.LEFT, padx=2)
        tk.Button(pc_buttons, text="Generate",
                  command=self.generate_process_conditions).pack(side=tk.LEFT, padx=2)
        tk.Button(pc_buttons, text="View",
                  command=lambda: self.view_data("Process Conditions")).pack(side=tk.LEFT, padx=2)

        # Instrument Index Row
        ii_frame = tk.Frame(main_frame)
        ii_frame.pack(fill=tk.X, pady=5)

        ii_label = tk.Label(ii_frame, text="Instrument Index", width=15)
        ii_label.pack(side=tk.LEFT, padx=5)

        ii_entry = tk.Entry(ii_frame)
        ii_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        self.entries.append((ii_entry, "tag_data"))

        ii_buttons = tk.Frame(ii_frame)
        ii_buttons.pack(side=tk.RIGHT)

        tk.Button(ii_buttons, text="Browse",
                  command=lambda: self.browse(ii_entry, "tag_data")).pack(side=tk.LEFT, padx=2)
        tk.Button(ii_buttons, text="Configure",
                  command=lambda: self.configure("Instrument Index")).pack(side=tk.LEFT, padx=2)
        tk.Button(ii_buttons, text="Generate",
                  command=self.generate_tag_data).pack(side=tk.LEFT, padx=2)
        tk.Button(ii_buttons, text="View",
                  command=lambda: self.view_data("Instrument Index")).pack(side=tk.LEFT, padx=2)

        # Datasheets Row
        ds_frame = tk.Frame(main_frame)
        ds_frame.pack(fill=tk.X, pady=5)

        ds_label = tk.Label(ds_frame, text="Datasheets", width=15)
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

    def configure_coordinate_value_data(self):
        if not self.app:
            self.app = xw.App(visible=True)

        self.wb = self.app.books.open(self.datasheet_path)
        # self.wb = xw.Book(self.datasheet_path)
        # the end goal here is to have a dict of pc_coordinate_values and td_coordinate_values each of which is just
        # has values like 'A1':'LINE', 'B2':'Max Pressure'
        td_combo_values = []
        # Example values for the combo box
        for key, value in self.td.items():
            td_combo_values = list(value.keys())
            print(list(td_combo_values))
            break

        pc_combo_values = []
        # Example values for the combo box
        for key, value in self.pc.items():
            pc_combo_values = list(value.keys())
            print(list(pc_combo_values))
            break

        def clear_td():
            self.td_coordinate_values = {}
            update_td_listbox()

        def clear_pc():
            self.pc_coordinate_values = {}
            update_pc_listbox()

        def add_to_td():
            key = key_entry.get()
            td_value = td_combo.get()
            # pc_value = pc_combo.get()

            if key and td_value:
                self.td_coordinate_values[key] = td_value
                print("Added:", key, td_value)
                # You can add further actions here or remove the print statement

        def add_to_pc():
            key = key_entry.get()
            # td_value = td_combo.get()
            pc_value = pc_combo.get()

            if key and pc_value:
                self.pc_coordinate_values[key] = pc_value
                print("Added:", key, pc_value)
                # You can add further actions here or remove the print statement

        configure_window = tk.Toplevel(self.root)
        configure_window.title("Configure Coordinate-Value Data")

        # Frame to hold the key label and entry widgets
        key_frame = ttk.Frame(configure_window)
        key_frame.pack(fill="x", padx=10, pady=5)  # Adjust padx and pady as needed

        # Key label
        key_label = ttk.Label(key_frame, text="Enter Key Coordinate:\n(First entry is top tag default)")
        key_label.pack(side="left")  # Pack label to the left within the frame

        # Key entry with width set to span the window
        entry_var = tk.StringVar()
        key_entry = ttk.Entry(key_frame, textvariable=entry_var)
        key_entry.pack(side="left", fill="x", expand=True)  # Pack entry to fill the available horizontal space

        # Frame to hold the key label and entry widgets
        listbox_frame = ttk.Frame(configure_window)
        listbox_frame.pack(fill="x", padx=10, pady=5)  # Adjust padx and pady as needed

        # Frame to hold the key label and entry widgets
        td_frame = ttk.Frame(listbox_frame)
        td_frame.pack(side="left", fill="x", expand=True, padx=5, pady=5)  # Adjust padx and pady as needed

        xfn_button = tk.Button(listbox_frame, text='Transformation\nCode and Key', command=self.get_xkey)
        xfn_button.pack(side='left')
        # Frame to hold the key label and entry widgets
        pc_frame = ttk.Frame(listbox_frame)
        pc_frame.pack(side="left", fill="x", expand=True, padx=5, pady=5)  # Adjust padx and pady as needed

        td_top_frame = ttk.Frame(td_frame)
        td_top_frame.pack(fill="x", expand=True, padx=5, pady=5)
        pc_top_frame = ttk.Frame(pc_frame)
        pc_top_frame.pack(fill="x", expand=True, padx=5, pady=5)

        # Combobox for td_combo_values
        td_label = ttk.Label(td_top_frame, text="Select TD Value:")
        td_label.pack(side='left')
        td_combo = ttk.Combobox(td_top_frame, values=td_combo_values, state="readonly")
        td_combo.pack(side="left", fill="x", expand=True, padx=5, pady=5)

        # Combobox for pc_combo_values
        pc_label = ttk.Label(pc_top_frame, text="Select PC Value:")
        pc_label.pack(side='left')
        pc_combo = ttk.Combobox(pc_top_frame, values=pc_combo_values, state="readonly")
        pc_combo.pack(side="left", fill="x", expand=True, padx=5, pady=5)

        td_listbox = tk.Listbox(td_frame)
        td_listbox.pack(fill="x", expand=True, padx=5, pady=5)

        pc_listbox = tk.Listbox(pc_frame)
        pc_listbox.pack(fill="x", expand=True, padx=5, pady=5)

        def update_td_listbox():
            # Clear previous items
            td_listbox.delete(0, tk.END)
            # Populate list box with items from self.td_coordinate_values
            for key, value in self.td_coordinate_values.items():
                td_listbox.insert(tk.END, f"{key}: {value}")

            first_entry = td_listbox.get(0)
            print(first_entry)
            try:
                self.top_tag = first_entry.split(':')[0]
            except:
                print('failed to set top_tag')

        def update_pc_listbox():
            # Clear previous items
            pc_listbox.delete(0, tk.END)
            # Populate list box with items from self.pc_coordinate_values
            for key, value in self.pc_coordinate_values.items():
                pc_listbox.insert(tk.END, f"{key}: {value}")

        # Function to update both list boxes
        def update_listboxes():
            update_td_listbox()
            update_pc_listbox()

        # Call update_listboxes initially to populate the list boxes
        update_listboxes()

        # Add buttons for TD and PC
        add_td_button = ttk.Button(td_frame, text="Add to TD",
                                   command=lambda: [add_to_td(), update_td_listbox()])
        add_td_button.pack(side='right', padx=10)

        add_pc_button = ttk.Button(pc_frame, text="Add to PC",
                                   command=lambda: [add_to_pc(), update_pc_listbox()])
        add_pc_button.pack(side='right', padx=10)

        # Function to remove item from self.td_coordinate_values
        def remove_from_td():
            selected_index = td_listbox.curselection()
            if selected_index:
                key_to_remove = list(self.td_coordinate_values.keys())[selected_index[0]]
                del self.td_coordinate_values[key_to_remove]
                update_td_listbox()

        # Function to remove item from self.pc_coordinate_values
        def remove_from_pc():
            selected_index = pc_listbox.curselection()
            if selected_index:
                key_to_remove = list(self.pc_coordinate_values.keys())[selected_index[0]]
                del self.pc_coordinate_values[key_to_remove]
                update_pc_listbox()

        # Remove from TD button
        remove_td_button = ttk.Button(td_frame, text="Remove from TD", command=remove_from_td)
        remove_td_button.pack(side='right', padx=10)

        # Remove from PC button
        remove_pc_button = ttk.Button(pc_frame, text="Remove from PC", command=remove_from_pc)
        remove_pc_button.pack(side='right', padx=10)

        # Remove from TD button
        clear_td_button = ttk.Button(td_frame, text="Clear All TD", command=clear_td)
        clear_td_button.pack(side='right', padx=10)

        # Remove from PC button
        clear_pc_button = ttk.Button(pc_frame, text="Clear All PC", command=clear_pc)
        clear_pc_button.pack(side='right', padx=10)

        def update_entry():
            current_selection = xw.apps.active.selection.address
            current_selection = current_selection.split(':')[0]
            current_selection = current_selection.replace('$', '')
            entry_var.set(current_selection)
            root.after(200, update_entry)

        configure_window.after(200, update_entry)

        def on_configure_window_close():
            try:
                app.quit()
            except Exception as e:
                print(f"Error closing Excel: {e}")
            finally:
                configure_window.destroy()

        # Add this to your configure_window setup
        # configure_window.protocol("WM_DELETE_WINDOW", on_configure_window_close)

    def create_coordinates_tab(self, tab):
        def init_excel():
            if not self.app:
                self.app = xw.App(visible=True)
            if self.datasheet_path:
                self.wb = self.app.books.open(self.datasheet_path)

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
                td_listbox.insert(tk.END, f"{key}: {value}")
            try:
                first_entry = td_listbox.get(0)
                self.top_tag = first_entry.split(':')[0]
            except:
                print('failed to set top_tag')

        def update_pc_listbox():
            pc_listbox.delete(0, tk.END)
            for key, value in self.pc_coordinate_values.items():
                pc_listbox.insert(tk.END, f"{key}: {value}")

        def update_listboxes():
            update_td_listbox()
            update_pc_listbox()

        def reinitialize():
            init_excel()
            td_values, pc_values = get_combo_values()
            td_combo['values'] = td_values
            pc_combo['values'] = pc_values
            update_listboxes()

        # Initial Excel setup if path exists
        init_excel()
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

        xfn_button = ttk.Button(content_frame, text="Transformation\nCode and Key", command=self.get_xkey)
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
            ("Source Sheet Name", "source_sheet_name", ttk.Combobox),
            ("Datasheet Coordinate", "datasheet_coord", ttk.Entry),
            ("Datasheet String", "ds_str", ttk.Entry),
            ("Tag Pattern", "tag_pattern", ttk.Entry),
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

        def save_datasheet_settings():
            for attr, entry in entries.items():
                if attr == 'rows_per_sheet':
                    try:
                        setattr(self, attr, int(entry.get()))
                    except ValueError:
                        setattr(self, attr, 0)
                else:
                    setattr(self, attr, entry.get())

        save_btn = ttk.Button(tab, text="Save", command=save_datasheet_settings)
        save_btn.grid(row=len(fields), column=0, columnspan=2, pady=10)
        tab.grid_columnconfigure(1, weight=1)

    def create_filters_tab(self, tab):
        # Add filter frame
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


    def get_sheet_names(self):
        if not self.datasheet_path:
            return []
        try:
            wb = openpyxl.load_workbook(self.datasheet_path, read_only=True)
            return wb.sheetnames
        except Exception as e:
            print(f"Error getting sheet names: {e}")
            return []

    def update_pc_keys(self):
        code = askstring("Enter transformation code", "Enter transformation code", initialvalue='"-".join(x.split("-")[-2:])')
        self.pc = transform_dictionary(self.pc, code)

    def delete_added_sheets(self):
        for sheet in self.new_sheets:
            self.wb.sheets[sheet].delete()


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

    def configure_add_datasheet(self):
        configure_window = tk.Toplevel()
        configure_window.title("Configure Add Datasheet")

        # Function to update attributes with values from Entry widgets
        def update_attributes():
            self.source_sheet_name = source_sheet_name_entry.get()
            # Convert string input to tuple for datasheet coordinates
            self.datasheet_coord = datasheet_coord_entry.get()

            self.ds_str = ds_str_entry.get()
            self.tag_pattern = tag_pattern_entry.get()
            self.top_tag = top_tag_entry.get()
            # Convert string input to integer for rows per sheet
            try:
                self.rows_per_sheet = int(rows_per_sheet_entry.get())
            except ValueError:
                # Handle incorrect input gracefully (e.g., default to 0)
                self.rows_per_sheet = 0

            # Close the configuration window after updating attributes
            configure_window.destroy()

        sn = []
        try:
            wb = openpyxl.load_workbook(self.datasheet_path, read_only=True)
            sn = wb.sheetnames
        except Exception as e:
            print(e)

        tk.Label(configure_window, text="Source Sheet Name:").pack(anchor='sw', pady=(10, 2))
        source_sheet_name_entry = ttk.Combobox(configure_window, values=sn)
        source_sheet_name_entry.insert(tk.END, self.source_sheet_name)
        source_sheet_name_entry.pack(fill='x')

        tk.Label(configure_window, text="Datasheet Coordinates:").pack(anchor='sw', pady=(10, 2))
        datasheet_coord_entry = tk.Entry(configure_window)
        datasheet_coord_entry.insert(tk.END, self.datasheet_coord)  # Display current row value
        datasheet_coord_entry.pack(fill='x')

        tk.Label(configure_window, text="Datasheet String:").pack(anchor='sw', pady=(10, 2))
        ds_str_entry = tk.Entry(configure_window)
        ds_str_entry.insert(tk.END, self.ds_str)
        ds_str_entry.pack(fill='x')

        tk.Label(configure_window, text="Tag Pattern (Regex):").pack(anchor='sw', pady=(10, 2))
        tag_pattern_entry = tk.Entry(configure_window)
        tag_pattern_entry.insert(tk.END, self.tag_pattern)
        tag_pattern_entry.pack(fill='x')

        tk.Label(configure_window, text="Top Tag Coordinate:").pack(anchor='sw', pady=(10, 2))
        top_tag_entry = tk.Entry(configure_window)
        top_tag_entry.insert(tk.END, str(self.top_tag))  # Display current rows per sheet value
        top_tag_entry.pack(fill='x')

        tk.Label(configure_window, text="Rows per Sheet:").pack(anchor='sw', pady=(10, 2))
        rows_per_sheet_entry = tk.Entry(configure_window)
        rows_per_sheet_entry.insert(tk.END, str(self.rows_per_sheet))  # Display current rows per sheet value
        rows_per_sheet_entry.pack(fill='x')

        # Button to update attributes
        tk.Button(configure_window, text="Save", command=update_attributes).pack(pady=(10, 2))

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

    def open_excel_search_app(self):
        # Instantiate and show the ExcelSearchApp
        excel_search_app = ExcelSearchApp()
        excel_search_app.mainloop()

    def open_edit_xlsx(self):
        edit_xlsx_window = tk.Toplevel(self.root)
        ExcelEditorApp(edit_xlsx_window)

    def generate_process_conditions(self):
        # print("Generating Process Conditions. Make sure there isn't a blank row between header and first entry!")
        self.pc = generate_dictionary_from_xlsx(self.process_conditions_path, self.pc_headers)
        show_nested_dict_analysis(self.pc)
        print("Generated Process Conditions")

    def generate_tag_data(self):
        print("Generating Tag Data")
        self.td = generate_dictionary_from_xlsx(self.tag_data_path, self.td_headers)
        show_nested_dict_analysis(self.td)
        print("Generated Tag Data")

    def assign_value_coordinate_to_tag(self):
        print("Generating Coordinate-Value Data")
        self.tag_cell_values = {}  # 'a1':'LINE', 'a2':'PID' ...
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

                # Rest of the function remains the same
                data = {}
                for coordinate, value in self.td_coordinate_values.items():
                    print(f'tag {tag}, value {value}, coord {coordinate}')
                    data[coordinate] = self.td[tag][value]

                try:
                    interface = translate(self.td[tag][self.td_xkey], self.transformation_code)
                    print(f'tag: {tag}, td_xkey: {self.td_xkey}, interface: {interface}')
                    print("length ", self.pc_coordinate_values)

                    for coordinate, value in self.pc_coordinate_values.items():
                        print("value ", value)
                        try:
                            data[coordinate] = self.pc[interface][value]
                        except Exception as e:
                            print(e)

                except Exception as e:
                    print('xkey pc interface fail: ', )

                self.tag_cell_values[tag] = data

        print("Coordinate Values generated:", self.tag_cell_values)

    def add_datasheets(self):
        print('assigining tag coordinates')
        self.assign_value_coordinate_to_tag()
        print("Adding Datasheets")
        print(f'datasheet_path: {self.datasheet_path}')
        print(f'source_sheet_name: {self.source_sheet_name}')
        print(f'tag_cell_values: {self.tag_cell_values}')
        print(f'datasheet_coord: {self.datasheet_coord}')
        print(f'ds_str: {self.ds_str}')
        print(f'tag_pattern: {self.tag_pattern}')
        print(f'rows_per_sheet: {self.rows_per_sheet}')
        # add_datasheets2 has a key coordinate add_datasheets doesn't

        if not self.app:
            self.app = xw.App(visible=True)
        self.wb = self.app.books.open(self.datasheet_path)
        print('opened ',self.datasheet_path)
        print('sheet', self.wb.sheets[self.source_sheet_name])
        self.new_sheets = add_datasheets(self.wb, self.source_sheet_name, self.tag_cell_values, self.datasheet_coord,
                        self.ds_str, rows_per_sheet=self.rows_per_sheet, key_coordinate=self.top_tag)

        # self.wb.close()
        print("DONE")

    def save_and_close_workbook(self):
        """
        Save and close an xlwings workbook

        Args:
            wb: xlwings Workbook object
            filepath: Optional path to save the file. If None, saves in current location
        """
        try:

            self.wb.save()
            self.wb.close()
            self.wb = None
            print('ran save an close')
        except Exception as e:
            print('error saving workbook ', e)

    def update_entry(self, entry, variable):
        filename = ''
        if variable == "process_conditions":
            filename = self.process_conditions_path
        elif variable == "tag_data":
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

        if variable == "process_conditions":
            self.process_conditions_path = filename
        elif variable == "tag_data":
            self.tag_data_path = filename
        elif variable == "coordinate_value":
            self.coordinate_value_path = filename
        elif variable == "datasheets":
            self.datasheet_path = filename
            print(self.datasheet_path)

    def configure(self, text):
        print(f"Configure {text}")

        if text == 'Instrument Index':
            self.configure_tag_data()

        if text == 'Process Conditions':
            self.configure_process_conditions()

        if text == 'Coordinate-Value Data':
            self.configure_ds()

        if text == 'Datasheets':
            self.configure_ds()

    def configure_tag_data(self):
        self.configure_list("Tag Data Headers", self.td_headers)

    def configure_process_conditions(self):
        self.configure_list("Process Conditions Headers", self.pc_headers)

    def configure_list(self, desc, headers):
        configure_window = tk.Toplevel(self.root)
        configure_window.title(desc)
        configure_window.geometry("320x320")

        # Create a copy of the original list
        original_headers = headers.copy()

        # Function to add an item to headers
        def add_item():
            new_item = entry.get()
            if new_item:
                original_headers.append(new_item)
                update_listbox()

        # Function to remove the selected item from headers
        def remove_item():
            selected_index = listbox.curselection()
            if selected_index:
                index = selected_index[0]
                del original_headers[index]
                update_listbox()

        # Function to update the listbox with headers
        def update_listbox():
            listbox.delete(0, tk.END)
            for item in original_headers:
                listbox.insert(tk.END, item)

        # Function to save changes to the original list
        def save_changes(headers):
            headers.clear()  # Clear the original headers list
            headers.extend(original_headers)  # Update the original list with modified headers
            configure_window.destroy()  # Close the window after saving changes

        # Creating GUI elements
        entry_frame = tk.Frame(configure_window)
        entry_frame.pack(fill=tk.X)

        entry = tk.Entry(entry_frame)
        entry.pack(side=tk.LEFT, expand=True, fill=tk.X)

        add_button = tk.Button(entry_frame, text="Add Item", command=add_item)
        add_button.pack(side=tk.RIGHT)

        listbox = tk.Listbox(configure_window)
        listbox.pack(fill=tk.BOTH, expand=True)

        remove_button = tk.Button(configure_window, text="Remove Item", command=remove_item)
        remove_button.pack(fill=tk.X)

        save_button = tk.Button(configure_window, text="Save", command=lambda: save_changes(headers))
        save_button.pack(fill=tk.X)

        # Update the listbox initially if headers has items
        update_listbox()

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

                # Force tab content refresh
                self.refresh_tab_content()
                self.update_entries()
                print(f"Settings loaded successfully from {file_ext} file!")

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

                # Force tab content refresh
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

    def update_datasheet(self):
        update_datasheets(self.datasheet_path, self.tag_cell_values, self.ds_str, self.rows_per_sheet,
                          key_coordinate=self.top_tag)

    def configure_ds(self):
        self.refresh_tab_content()
        self.update_entries()


if __name__ == "__main__":
    root = tk.Tk()
    app = DatasheetGeneratorApp(root)
    root.mainloop()
