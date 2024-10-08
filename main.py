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
class DataGeneratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Data Generator App")
        self.create_widgets()
        self.workbook = None
        self.parameters = {

            'process_conditions_path': '',
            'tag_data_path': '',
            'td_coordinate_values': {},
            'pc_coordinate_values': {},
            'td': {},
            'td_headers': ['TAG NUMBER'],
            'pc': {},
            'pc_headers': ['Line No.'],
            'transformation_code': 'x.split("-")[2]',
            'td_xkey': '',
            'tag_filters': [],
            'tag_cell_values': {},
            'datasheet_path': '',
            'source_sheet_name': 'TEMPLATE',
            'datasheet_coord': 'U8',
            'ds_str': 'DS-IA-',
            'tag_pattern': r'^[0-9]{3}-[A-Z]{2,3}-[0-9]{4}[A-Z]?$',
            'rows_per_sheet': 1,
            'top_tag':'A1'
        }

        # Initialize all parameters in the __init__ method
        for param, value in self.parameters.items():
            setattr(self, param, value)
    def create_widgets(self):

        self.root.columnconfigure(2, weight=1)

        self.menu_bar = tk.Menu(self.root)
        self.root.config(menu=self.menu_bar)

        self.command_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label="Commands", menu=self.command_menu)

        self.command_menu.add_command(label="Load Settings", command=self.load_settings)
        self.command_menu.add_command(label="Save Settings", command=self.save_settings)
        self.command_menu.add_command(label="Save Settings except", command=self.save_settings_except)
        self.command_menu.add_command(label="Transformation Code and Key", command=self.get_xkey)
        self.command_menu.add_command(label="Index Filters", command=self.set_tag_filters)
        self.command_menu.add_command(label="Populate Index (td) from Json", command=self.load_td_from_json)
        self.command_menu.add_command(label="Populate Index (td) from Datasheet", command=self.load_td_from_datasheet)
        self.command_menu.add_command(label="Populate Index (pc) from Datasheet", command=self.load_pc_from_datasheet)
        self.command_menu.add_command(label="Update Datasheet", command=self.update_datasheet)
        self.command_menu.add_command(label="Run xlsx search app", command=self.open_excel_search_app)
        self.command_menu.add_command(label="Save and close", command=self.save_and_close_workbook)
        self.command_menu.add_command(label="Populate Headers", command=self.open_edit_xlsx)

        # filters_button = tk.Button(top_row_frame, text="Sorting Function")#
        #self.command_menu.add_command(label="Sorting Function", command=self.sorting_function)  # Assuming you have a sorting function defined

        # Increment row counter for subsequent rows
        idx = 1

        # Previous code for label texts, entries, and buttons
        self.label_texts = ["Process Conditions", "Instrument Index", "Coordinate-Value Data", "Datasheets"]
        label_texts = self.label_texts

        button_actions = [self.generate_process_conditions, self.generate_tag_data,
                          self.assign_value_coordinate_to_tag, self.add_datasheets]
        browse_variables = ["process_conditions", "tag_data", "coordinate_value", "datasheets"]

        self.entries = []
        for (label_text, action, browse_var) in zip(label_texts, button_actions, browse_variables):


            label = tk.Label(self.root, text=label_text)
            label.grid(row=idx, column=0, pady=5)

            # Help Button
            help_button = tk.Button(self.root, text="?", command=lambda text=label_text: self.show_help(text))
            help_button.grid(row=idx, column=1, padx=5)

            entry = tk.Entry(self.root)
            self.entries.append((entry, browse_var))# used for prepopulating entry boxes
            entry.grid(row=idx, column=2, padx=5, sticky='ew')

            browse_button = tk.Button(self.root, text="Browse",
                                      command=lambda entry=entry, var=browse_var: self.browse(entry, var))
            browse_button.grid(row=idx, column=3, padx=5)

            configure_button = tk.Button(self.root, text="Configure", command=lambda text=label_text: self.configure(text))
            configure_button.grid(row=idx, column=4, padx=5)

            button = tk.Button(self.root, text="Generate", command=action)
            button.grid(row=idx, column=5, padx=5)

            view_button = tk.Button(self.root, text="View", command=lambda text=label_text: self.view_data(text))
            view_button.grid(row=idx, column=6, padx=5)
            idx +=1

    def open_excel_search_app(self):
        # Instantiate and show the ExcelSearchApp
        excel_search_app = ExcelSearchApp()
        excel_search_app.mainloop()

    def open_edit_xlsx(self):
        edit_xlsx_window = tk.Toplevel(self.root)
        ExcelEditorApp(edit_xlsx_window)

    def generate_process_conditions(self):
        #print("Generating Process Conditions. Make sure there isn't a blank row between header and first entry!")
        self.pc = generate_dictionary_from_xlsx(self.process_conditions_path, self.pc_headers)
        print("Generated Process Conditions")

    def generate_tag_data(self):
        print("Generating Tag Data")
        self.td = generate_dictionary_from_xlsx(self.tag_data_path, self.td_headers)
        print("Generated Tag Data")

    def translate(self, input_string):
        try:
            x_fn = eval(f'lambda x: {self.transformation_code}')
            transformed_string = x_fn(input_string)
            return transformed_string
        except Exception as e:
            return f'error applying transformation: {e}'
        # modify string
        # for example result = data_string.split('-')[2]

    def update_key_entry(self):
        # Get the initial selection
        initial_selection = xw.apps.active.selection.address

        while True:
            # Get the current selection
            current_selection = xw.apps.active.selection.address

            # Check if a new cell is selected
            if current_selection != initial_selection:
                # Update initial_selection to the current selection
                initial_selection = current_selection

                # Get the active sheet
                sheet = self.wb.sheets.active

                # Get the value of the newly selected cell
                selected_cell_value = sheet.range(current_selection).value

                # Update key_entry with the coordinate
                self.key_entry = current_selection
                print(f"Selected cell value: {selected_cell_value}, Key entry: {self.key_entry}")

    def configure_coordinate_value_data(self):

        self.wb = xw.Book(self.datasheet_path)
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
            #pc_value = pc_combo.get()

            if key and td_value:
                self.td_coordinate_values[key] = td_value
                print("Added:", key, td_value)
                # You can add further actions here or remove the print statement
        def add_to_pc():
            key = key_entry.get()
            #td_value = td_combo.get()
            pc_value = pc_combo.get()

            if key and pc_value:
                self.pc_coordinate_values[key] = pc_value
                print("Added:", key, pc_value)
                # You can add further actions here or remove the print statement

        def increment():
            value = key_entry.get()  # Assuming key_entry is the tkinter Entry widget

            pattern = r'^([a-zA-Z]+)(\d+)$'  # Regular expression pattern to match 'alpha' + 'numeric'
            match = re.match(pattern, value)

            if match:
                alpha_part = match.group(1)
                numeric_part = int(match.group(2))
                incremented_numeric = numeric_part + 1
                new_value = f"{alpha_part}{incremented_numeric}"
                key_entry.delete(0, 'end')  # Clear the Entry widget
                key_entry.insert(0, new_value)  # Update the Entry widget with the new value

        def decrement():
            value = key_entry.get()  # Assuming key_entry is the tkinter Entry widget

            pattern = r'^([a-zA-Z]+)(\d+)$'  # Regular expression pattern to match 'alpha' + 'numeric'
            match = re.match(pattern, value)

            if match:
                alpha_part = match.group(1)
                numeric_part = int(match.group(2))

                # Check if the numeric part is greater than 1 before decrementing
                if numeric_part > 1:
                    decremented_numeric = numeric_part - 1
                    new_value = f"{alpha_part}{decremented_numeric}"
                    key_entry.delete(0, 'end')  # Clear the Entry widget
                    key_entry.insert(0, new_value)  # Update the Entry widget with the new value

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

        # Key entry with width set to span the window
        inc_button = tk.Button(key_frame,text='+', command=increment, padx=5)
        inc_button.pack(side="left")  # Pack entry to fill the available horizontal space

        # Key entry with width set to span the window
        dec_button = tk.Button(key_frame,text='-', command=decrement, padx=5)
        dec_button.pack(side="left")  # Pack entry to fill the available horizontal space


        # Frame to hold the key label and entry widgets
        listbox_frame = ttk.Frame(configure_window)
        listbox_frame.pack(fill="x", padx=10, pady=5)  # Adjust padx and pady as needed

        # Frame to hold the key label and entry widgets
        td_frame = ttk.Frame(listbox_frame)
        td_frame.pack(side="left", fill="x",  expand=True, padx=5, pady=5)  # Adjust padx and pady as needed


        xfn_button = tk.Button(listbox_frame, text='Transformation\nCode and Key', command=self.get_xkey)
        xfn_button.pack(side='left')
        # Frame to hold the key label and entry widgets
        pc_frame = ttk.Frame(listbox_frame)
        pc_frame.pack(side="left", fill="x",  expand=True, padx=5, pady=5)  # Adjust padx and pady as needed

        td_top_frame = ttk.Frame(td_frame)
        td_top_frame.pack(fill="x",  expand=True, padx=5, pady=5)
        pc_top_frame = ttk.Frame(pc_frame)
        pc_top_frame.pack(fill="x",  expand=True, padx=5, pady=5)


        # Combobox for td_combo_values
        td_label = ttk.Label(td_top_frame, text="Select TD Value:")
        td_label.pack(side='left')
        td_combo = ttk.Combobox(td_top_frame, values=td_combo_values, state="readonly")
        td_combo.pack(side="left", fill="x",  expand=True, padx=5, pady=5)

        # Combobox for pc_combo_values
        pc_label = ttk.Label(pc_top_frame, text="Select PC Value:")
        pc_label.pack(side='left')
        pc_combo = ttk.Combobox(pc_top_frame, values=pc_combo_values, state="readonly")
        pc_combo.pack(side="left", fill="x",  expand=True, padx=5, pady=5)



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

    def get_xkey(self):
        xkey_window = tk.Toplevel(self.root)
        xkey_window.title("Configure translation_lambda")


        frame1 = tk.Frame(xkey_window)
        frame1.pack(fill="x", pady=5,padx=5)
        frame2 = tk.Frame(xkey_window, pady=5,padx=5)
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

    def assign_value_coordinate_to_tag(self):
        print("Generating Coordinate-Value Data")
        self.tag_cell_values = {}# 'a1':'LINE', 'a2':'PID' ...
        for tag in self.td:
            if tag:
                # filter out
                continue_flag = False
                for header, key in self.tag_filters:
                    print('tag',tag)
                    print('header',header)
                    if self.td[tag][header] != key:
                        continue_flag = True
                        # continue

                if continue_flag:
                    continue


                # this can be confusing but if you follow the logic here were just
                # giving the coordinates and actual value from the key
                data = {}
                for coordinate, value in self.td_coordinate_values.items():
                    #creating a new entry in data which will look like 'a1':'3"-SW-132-01001-B1A2-IH'
                    data[coordinate] = self.td[tag][value]

                # interface should end up with a key that fits in self.pc, so maybe a line number
                # translate takes in a datapoint from the td and extracts an item to match in
                try:
                    interface = self.translate(self.td[tag][self.td_xkey])
                    print(f'tag: {tag}, td_xkey: {self.td_xkey}, interface: {interface}')
                    print("length ", self.pc_coordinate_values)

                    for coordinate, value in self.pc_coordinate_values.items():
                        print("value ", value)
                        # creating a new entry in data which will look like 'a1':'3"-SW-132-01001-B1A2-IH'
                        try:
                            data[coordinate] = self.pc[interface][value]
                        except Exception as e:
                            print(e)

                except Exception as e:
                    print('xkey pc interface fail: ', )




                self.tag_cell_values[tag] = data

        #self.tag_cell_values = tag_cell_values


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
        #add_datasheets2 has a key coordinate add_datasheets doesn't


        self.datasheet = add_datasheets2(self.datasheet_path, self.source_sheet_name, self.tag_cell_values, self.datasheet_coord,
                       self.ds_str, rows_per_sheet=self.rows_per_sheet, key_coordinate=self.top_tag)
        print("DONE")

    def save_and_close_workbook(self):
        """
        Save and close an xlwings workbook

        Args:
            wb: xlwings Workbook object
            filepath: Optional path to save the file. If None, saves in current location
        """
        try:
            self.workbook.save()
            self.workbook.close()
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

        tk.Label(configure_window, text="Source Sheet Name:").pack(anchor='sw', pady=(10,2))
        source_sheet_name_entry = ttk.Combobox(configure_window, values=sn)
        source_sheet_name_entry.insert(tk.END, self.source_sheet_name)
        source_sheet_name_entry.pack(fill='x')

        tk.Label(configure_window, text="Datasheet Coordinates:").pack(anchor='sw', pady=(10,2))
        datasheet_coord_entry = tk.Entry(configure_window)
        datasheet_coord_entry.insert(tk.END, self.datasheet_coord)  # Display current row value
        datasheet_coord_entry.pack(fill='x')

        tk.Label(configure_window, text="Datasheet String:").pack(anchor='sw', pady=(10,2))
        ds_str_entry = tk.Entry(configure_window)
        ds_str_entry.insert(tk.END, self.ds_str)
        ds_str_entry.pack(fill='x')

        tk.Label(configure_window, text="Tag Pattern (Regex):").pack(anchor='sw', pady=(10,2))
        tag_pattern_entry = tk.Entry(configure_window)
        tag_pattern_entry.insert(tk.END, self.tag_pattern)
        tag_pattern_entry.pack(fill='x')

        tk.Label(configure_window, text="Top Tag Coordinate:").pack(anchor='sw', pady=(10,2))
        top_tag_entry = tk.Entry(configure_window)
        top_tag_entry.insert(tk.END, str(self.top_tag))  # Display current rows per sheet value
        top_tag_entry.pack(fill='x')

        tk.Label(configure_window, text="Rows per Sheet:").pack(anchor='sw', pady=(10,2))
        rows_per_sheet_entry = tk.Entry(configure_window)
        rows_per_sheet_entry.insert(tk.END, str(self.rows_per_sheet))  # Display current rows per sheet value
        rows_per_sheet_entry.pack(fill='x')

        # Button to update attributes
        tk.Button(configure_window, text="Save", command=update_attributes).pack(pady=(10,2))

    def configure(self, text):
        print(f"Configure {text}")

        if text == 'Instrument Index':
            self.configure_tag_data()

        if text == 'Process Conditions':
            self.configure_process_conditions()

        if text == 'Coordinate-Value Data':
            self.configure_coordinate_value_data()

        if text == 'Datasheets':
            self.configure_add_datasheet()

    def configure_tag_data(self):
        self.configure_list("Tag Data Headers",self.td_headers)

    def configure_process_conditions(self):
        self.configure_list("Process Conditions Headers",self.pc_headers)

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

    def set_tag_filters(self):
        view_window = tk.Toplevel(self.root)
        view_window.title("Set Tag Filters. ReGenerate Coordinates if necessary")

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
            name_entry =  ttk.Combobox(view_window, values=td_combo_values)
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

    def load_settings(self):
        file_path = filedialog.askopenfilename(filetypes=[("JSON files", "*.json")])
        if file_path:
            try:
                with open(file_path, 'r') as file:
                    settings_data = json.load(file)
                    # Update the parameters with loaded settings data
                    for key, value in settings_data.items():
                        if key in self.parameters:
                            setattr(self, key, value)
                    self.update_entries()
                    print("Settings loaded successfully!")

            except FileNotFoundError:
                print("File not found. Unable to load settings.")
            except json.JSONDecodeError:
                print("Error decoding JSON. Unable to load settings.")

    def save_settings(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON files", "*.json")])
        if file_path:
            settings_to_save = {key: getattr(self, key) for key in self.parameters}
            try:
                with open(file_path, 'w') as file:
                    json.dump(settings_to_save, file, indent=4)
                    print("Settings saved successfully!")
            except Exception as e:
                print(f"Error occurred while saving settings: {e}")

    def save_settings_except(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON files", "*.json")])
        if file_path:
            # Use a set to efficiently check for exclusions
            exclude_set = ('td', 'pc')
            settings_to_save = {key: getattr(self, key) for key in dir(self) if hasattr(self, key) and key not in exclude_set}
            try:
                with open(file_path, 'w') as file:
                    json.dump(settings_to_save, file, indent=4)
                    print("Settings saved successfully!")
            except Exception as e:
                print(f"Error occurred while saving settings: {e}")

    def show_help(self, text):
        help_messages = {
            "Process Conditions":"""
    Generates process conditions by extracting data from an Excel workbook.
    
    This method initiates the retrieval of process conditions from an Excel workbook 
    located at the specified path. It reads the workbook using openpyxl, processes 
    the data from multiple worksheets, and generates a consolidated dictionary of 
    process conditions.

    Steps:
    1. Opens the Excel workbook located at the defined path.
    2. Retrieves process conditions from multiple worksheets within the workbook:
        - Iterates through each worksheet in the workbook.
        - For each worksheet, identifies and extracts specific process condition data 
          based on predefined header values.
        - Utilizes helper functions to process the worksheet data, locate relevant 
          headers, and transform the tabular data into a structured dictionary format.
        - Collates the extracted process condition data from all worksheets into a 
          consolidated dictionary, ensuring uniqueness of keys and combining similar 
          process conditions.
    3. Consolidates the extracted process conditions into a single dictionary.
    
    Note:
    - The method relies on helper functions for specific data extraction tasks.
    - Assumes that the workbook's process conditions are structured in accordance 
      with the defined headers.

    Parameters:
    - self: Instance of the class where this method is implemented.

    Returns:
    - None: The method doesn't return any values but updates the class attribute 'pc'
      with the generated process conditions dictionary.

    Raises:
    - Any exceptions raised during the process are propagated upwards for handling.
    """,


            "Instrument Index": '''## Function: generate_tag_dictionary

### Description:
This function generates a dictionary of tags mapped to their respective rows from an Excel worksheet. It utilizes the `table_to_dict` helper function to convert the worksheet into a list of dictionaries and then creates a dictionary where the keys are unique tags and the values are the rows associated with each tag.

### Parameters:
- `path` (str): The file path of the Excel workbook.
- `sheet_name` (str): The name of the Excel worksheet.
- `start_row` (int): The row number where the data begins in the worksheet.
- `tag_header` (str): The header indicating the column containing tags.

### Returns:
- `tag_dictionary` (dict): A dictionary where keys are unique tags and values are rows associated with each tag.''',


            "Coordinate-Value Data":     """
    Generates coordinate-value data based on provided tag data and filters.

    This method iterates through the tag data (`self.td`) and applies filters (`self.tag_filters`) to extract relevant information.
    It creates a dictionary `self.tag_cell_values` where each tag serves as a key, containing a nested dictionary with coordinate-value pairs.

    Returns:
    None

    Steps:
    1. Initializes an empty dictionary `self.tag_cell_values` to store tag-specific coordinate-value pairs.
    2. Iterates through each tag in the tag data (`self.td`).
    3. Skips processing if the tag is empty (`if tag:`).
    4. Checks if the tag meets filter conditions specified in `self.tag_filters`.
        - If the tag does not satisfy the filter conditions, it skips to the next tag.
    5. Creates a new dictionary `data` to store coordinate-value pairs specific to the current tag.
    6. Iterates through the existing `coordinate_values` dictionary to map coordinates to their corresponding values from `self.td`.
    7. Stores the extracted coordinate-value pairs in the `data` dictionary.
    8. Associates the `data` dictionary with the current tag in `self.tag_cell_values`.
    9. Prints the generated coordinate values for each tag.

    Note:
    - `self.td`: Contains tag-specific information.
    - `self.tag_filters`: Specifies conditions to filter tags based on headers and keys.
    - `self.coordinate_values`: Contains coordinates with corresponding values to extract from `self.td`.

    Usage:
    - Call this method to generate tag-specific coordinate-value data based on the provided tag data and filters.
""",


            "Datasheets":     """
    Adds data from tag_cell_values dictionary to specific datasheets within an Excel file.

    This function populates datasheets in an Excel workbook based on provided parameters and data from the 'tag_cell_values' dictionary.

    Args:
    - datasheet_path (str): Path to the Excel file where datasheets will be modified.
    - source_sheet_name (str): Name of the source sheet in the Excel file.
    - tag_cell_values (dict): Dictionary containing tag-specific coordinate-value pairs.
    - datasheet_coord (str): Excel cell coordinates specifying where to insert the sheet number or identifier.
    - ds_str (str): String to be concatenated with the sheet number or identifier.
    - tag_pattern (str, optional): Regular expression pattern to validate tag format. Defaults to r'^[0-9]{3}-[A-Z]{2,3}-[0-9]{4}[A-Z]?$'.
    - rows_per_sheet (int, optional): Number of rows per sheet. Defaults to 1.
    - custom_sort (function, optional): Custom sorting function for keys in 'tag_cell_values'. Defaults to None.

    Returns:
    None

    Steps:
    1. Retrieves the count and existing tag set using the 'get_count' function based on 'rows_per_sheet' and 'tag_pattern'.
    2. Opens the specified datasheet.
    3. Gets the source sheet within the datasheet workbook.
    4. Sorts the keys in 'tag_cell_values' based on the provided 'custom_sort' function or default sorting.
    5. Iterates through sorted keys and adds data to respective datasheets based on conditions.
    6. Modifies existing sheets or creates new sheets as needed and populates them with cell values from 'tag_cell_values'.

    Note:
    - 'get_count' is a function that retrieves the count and existing tag set based on specified parameters.
    - 'tag_cell_values' is a dictionary containing tag-specific coordinate-value pairs.
    - The function handles adding data to Excel datasheets based on specified conditions.

    Usage:
    - Call this function to populate datasheets in an Excel workbook with data from the 'tag_cell_values' dictionary.

    Example Usage:
    ```
    add_datasheets('example.xlsx', 'SourceSheet', tag_cell_values_dict, 'A1', 'DS_', rows_per_sheet=2)
    ```

    """
        }

        help_text = help_messages.get(text, "No help available.")

        popup = tk.Tk()
        popup.title(f"{text} help")

        scrolled_text = scrolledtext.ScrolledText(popup, wrap=tk.WORD, width=40, height=10)
        scrolled_text.insert(tk.END, help_text)
        scrolled_text.config(state='disabled')
        scrolled_text.pack(expand=True, fill='both')

        popup.mainloop()



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
        #use tkinter to ask for the json file path
        #use load_dict_from_json(file_path) to set self.td
        # Ask the user to select a JSON file
        file_path = filedialog.askopenfilename(title="Select JSON file", filetypes=[("JSON files", "*.json")])

        # Check if a file was selected
        if file_path:
            # Load the JSON file using load_dict_from_json
            self.td = load_dict_from_json(file_path)

    def update_datasheet(self):
        update_datasheets(self.datasheet_path,self.tag_cell_values,self.ds_str,self.rows_per_sheet,
                          key_coordinate=self.top_tag)

if __name__ == "__main__":
    root = tk.Tk()
    app = DataGeneratorApp(root)
    root.mainloop()
