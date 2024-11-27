import tkinter as tk
from tkinter import ttk
import xlwings as xw
import re

class CoordinateValueConfigurator:
    def __init__(self, root, datasheet_path, td, pc):
        self.root = root
        self.datasheet_path = datasheet_path
        self.td = td
        self.pc = pc
        self.td_coordinate_values = {}
        self.pc_coordinate_values = {}
        self.top_tag = None

    def __call__(self):
        self.configure_coordinate_value_data()

    def configure_coordinate_value_data(self):
        self.wb = xw.Book(self.datasheet_path)

        td_combo_values = []
        pc_combo_values = []

        for key, value in self.td.items():
            td_combo_values = list(value.keys())
            break

        for key, value in self.pc.items():
            pc_combo_values = list(value.keys())
            break

        configure_window = tk.Toplevel(self.root)
        configure_window.title("Configure Coordinate-Value Data")

        # Key frame
        key_frame = ttk.Frame(configure_window)
        key_frame.pack(fill="x", padx=10, pady=5)

        key_label = ttk.Label(key_frame, text="Enter Key Coordinate:\n(First entry is top tag default)")
        key_label.pack(side="left")

        entry_var = tk.StringVar()
        key_entry = ttk.Entry(key_frame, textvariable=entry_var)
        key_entry.pack(side="left", fill="x", expand=True)

        inc_button = tk.Button(key_frame, text='+', command=lambda: self.increment(key_entry), padx=5)
        inc_button.pack(side="left")

        dec_button = tk.Button(key_frame, text='-', command=lambda: self.decrement(key_entry), padx=5)
        dec_button.pack(side="left")

        # Listbox frame
        listbox_frame = ttk.Frame(configure_window)
        listbox_frame.pack(fill="x", padx=10, pady=5)

        td_frame = ttk.Frame(listbox_frame)
        td_frame.pack(side="left", fill="x", expand=True, padx=5, pady=5)

        xfn_button = tk.Button(listbox_frame, text='Transformation\nCode and Key', command=self.get_xkey)
        xfn_button.pack(side='left')

        pc_frame = ttk.Frame(listbox_frame)
        pc_frame.pack(side="left", fill="x", expand=True, padx=5, pady=5)

        # TD and PC combo boxes
        td_top_frame = ttk.Frame(td_frame)
        td_top_frame.pack(fill="x", expand=True, padx=5, pady=5)

        pc_top_frame = ttk.Frame(pc_frame)
        pc_top_frame.pack(fill="x", expand=True, padx=5, pady=5)

        td_label = ttk.Label(td_top_frame, text="Select TD Value:")
        td_label.pack(side='left')

        td_combo = ttk.Combobox(td_top_frame, values=td_combo_values, state="readonly")
        td_combo.pack(side="left", fill="x", expand=True, padx=5, pady=5)

        pc_label = ttk.Label(pc_top_frame, text="Select PC Value:")
        pc_label.pack(side='left')

        pc_combo = ttk.Combobox(pc_top_frame, values=pc_combo_values, state="readonly")
        pc_combo.pack(side="left", fill="x", expand=True, padx=5, pady=5)

        td_listbox = tk.Listbox(td_frame)
        td_listbox.pack(fill="x", expand=True, padx=5, pady=5)

        pc_listbox = tk.Listbox(pc_frame)
        pc_listbox.pack(fill="x", expand=True, padx=5, pady=5)

        # Add buttons
        add_td_button = ttk.Button(td_frame, text="Add to TD",
                                   command=lambda: [self.add_to_td(key_entry, td_combo), self.update_td_listbox(td_listbox)])
        add_td_button.pack(side='right', padx=10)

        add_pc_button = ttk.Button(pc_frame, text="Add to PC",
                                   command=lambda: [self.add_to_pc(key_entry, pc_combo), self.update_pc_listbox(pc_listbox)])
        add_pc_button.pack(side='right', padx=10)

        # Remove buttons
        remove_td_button = ttk.Button(td_frame, text="Remove from TD", command=lambda: self.remove_from_td(td_listbox))
        remove_td_button.pack(side='right', padx=10)

        remove_pc_button = ttk.Button(pc_frame, text="Remove from PC", command=lambda: self.remove_from_pc(pc_listbox))
        remove_pc_button.pack(side='right', padx=10)

        # Clear buttons
        clear_td_button = ttk.Button(td_frame, text="Clear All TD", command=lambda: self.clear_td(td_listbox))
        clear_td_button.pack(side='right', padx=10)

        clear_pc_button = ttk.Button(pc_frame, text="Clear All PC", command=lambda: self.clear_pc(pc_listbox))
        clear_pc_button.pack(side='right', padx=10)

        self.update_listboxes(td_listbox, pc_listbox)
        configure_window.after(200, lambda: self.update_entry(entry_var))

    # ... (include all the other methods like increment, decrement, add_to_td, add_to_pc, etc.)

    def get_xkey(self):
        # Implement this method as needed
        pass

    def update_entry(self, entry_var):
        current_selection = xw.apps.active.selection.address
        current_selection = current_selection.split(':')[0]
        current_selection = current_selection.replace('$', '')
        entry_var.set(current_selection)
        self.root.after(200, lambda: self.update_entry(entry_var))