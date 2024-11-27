import tkinter as tk
from tkinter import ttk
import re
import xlwings as xw

from dataclasses import dataclass

@dataclass
class FieldMapperResults:
    td_coordinate_values: dict
    pc_coordinate_values: dict
    top_tag: str

class DatasheetFieldMapper:
    def __init__(self, parent, td, pc, td_coordinate_values, pc_coordinate_values, top_tag):
        self.window = tk.Toplevel(parent)
        self.window.title("Configure Coordinate-Value Data")
        self.td = td
        self.pc = pc
        self.td_coordinate_values = td_coordinate_values.copy()
        self.pc_coordinate_values = pc_coordinate_values.copy()
        self.top_tag = top_tag
        self.results = None
        self.app = None
        self.wb = None

        self.create_widgets()
        self.window.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.window.transient(parent)
        self.window.grab_set()

    def on_closing(self):
        self.results = FieldMapperResults(
            td_coordinate_values=self.td_coordinate_values,
            pc_coordinate_values=self.pc_coordinate_values,
            top_tag=self.top_tag
        )
        self.window.destroy()

    def get_results(self):
        self.window.wait_window()
        return self.results
    def create_widgets(self):
        # Frame for key coordinate
        key_frame = ttk.Frame(self.window)
        key_frame.pack(fill="x", padx=10, pady=5)

        tk.Label(key_frame, text="Enter Key Coordinate:\n(First entry is top tag default)").pack(side="left")

        self.entry_var = tk.StringVar()
        self.key_entry = ttk.Entry(key_frame, textvariable=self.entry_var)
        self.key_entry.pack(side="left", fill="x", expand=True)

        tk.Button(key_frame, text='+', command=self.increment, padx=5).pack(side="left")
        tk.Button(key_frame, text='-', command=self.decrement, padx=5).pack(side="left")

        # Frame for listboxes
        listbox_frame = ttk.Frame(self.window)
        listbox_frame.pack(fill="x", padx=10, pady=5)

        # TD Frame setup
        self.setup_td_frame(listbox_frame)

        # PC Frame setup
        self.setup_pc_frame(listbox_frame)

        self.window.after(200, self.update_entry)

    def setup_td_frame(self, parent):
        td_frame = ttk.Frame(parent)
        td_frame.pack(side="left", fill="x", expand=True, padx=5, pady=5)

        td_top_frame = ttk.Frame(td_frame)
        td_top_frame.pack(fill="x", expand=True, padx=5, pady=5)

        tk.Label(td_top_frame, text="Select TD Value:").pack(side='left')

        td_combo_values = self._get_td_combo_values()
        self.td_combo = ttk.Combobox(td_top_frame, values=td_combo_values, state="readonly")
        self.td_combo.pack(side="left", fill="x", expand=True, padx=5, pady=5)

        self.td_listbox = tk.Listbox(td_frame)
        self.td_listbox.pack(fill="x", expand=True, padx=5, pady=5)

        self.setup_td_buttons(td_frame)

    def setup_pc_frame(self, parent):
        pc_frame = ttk.Frame(parent)
        pc_frame.pack(side="left", fill="x", expand=True, padx=5, pady=5)

        pc_top_frame = ttk.Frame(pc_frame)
        pc_top_frame.pack(fill="x", expand=True, padx=5, pady=5)

        tk.Label(pc_top_frame, text="Select PC Value:").pack(side='left')

        pc_combo_values = self._get_pc_combo_values()
        self.pc_combo = ttk.Combobox(pc_top_frame, values=pc_combo_values, state="readonly")
        self.pc_combo.pack(side="left", fill="x", expand=True, padx=5, pady=5)

        self.pc_listbox = tk.Listbox(pc_frame)
        self.pc_listbox.pack(fill="x", expand=True, padx=5, pady=5)

        self.setup_pc_buttons(pc_frame)

    def setup_td_buttons(self, frame):
        ttk.Button(frame, text="Add to TD", command=lambda: [self.add_to_td(), self.update_td_listbox()]).pack(
            side='right', padx=10)
        ttk.Button(frame, text="Remove from TD", command=self.remove_from_td).pack(side='right', padx=10)
        ttk.Button(frame, text="Clear All TD", command=self.clear_td).pack(side='right', padx=10)

    def setup_pc_buttons(self, frame):
        ttk.Button(frame, text="Add to PC", command=lambda: [self.add_to_pc(), self.update_pc_listbox()]).pack(
            side='right', padx=10)
        ttk.Button(frame, text="Remove from PC", command=self.remove_from_pc).pack(side='right', padx=10)
        ttk.Button(frame, text="Clear All PC", command=self.clear_pc).pack(side='right', padx=10)

    def _get_td_combo_values(self):
        for key, value in self.td.items():
            return list(value.keys())
        return []

    def _get_pc_combo_values(self):
        for key, value in self.pc.items():
            return list(value.keys())
        return []

    def increment(self):
        value = self.key_entry.get()
        pattern = r'^([a-zA-Z]+)(\d+)$'
        match = re.match(pattern, value)
        if match:
            alpha_part = match.group(1)
            numeric_part = int(match.group(2))
            self.key_entry.delete(0, 'end')
            self.key_entry.insert(0, f"{alpha_part}{numeric_part + 1}")

    def decrement(self):
        value = self.key_entry.get()
        pattern = r'^([a-zA-Z]+)(\d+)$'
        match = re.match(pattern, value)
        if match:
            alpha_part = match.group(1)
            numeric_part = int(match.group(2))
            if numeric_part > 1:
                self.key_entry.delete(0, 'end')
                self.key_entry.insert(0, f"{alpha_part}{numeric_part - 1}")

    def update_entry(self):
        try:
            current_selection = xw.apps.active.selection.address
            current_selection = current_selection.split(':')[0].replace('$', '')
            self.entry_var.set(current_selection)
        except:
            pass
        self.window.after(200, self.update_entry)

    def add_to_td(self):
        key = self.key_entry.get()
        td_value = self.td_combo.get()
        if key and td_value:
            self.td_coordinate_values[key] = td_value
            self.update_top_tag()

    def add_to_pc(self):
        key = self.key_entry.get()
        pc_value = self.pc_combo.get()
        if key and pc_value:
            self.pc_coordinate_values[key] = pc_value

    def update_td_listbox(self):
        self.td_listbox.delete(0, tk.END)
        for key, value in self.td_coordinate_values.items():
            self.td_listbox.insert(tk.END, f"{key}: {value}")

    def update_pc_listbox(self):
        self.pc_listbox.delete(0, tk.END)
        for key, value in self.pc_coordinate_values.items():
            self.pc_listbox.insert(tk.END, f"{key}: {value}")

    def remove_from_td(self):
        selected_index = self.td_listbox.curselection()
        if selected_index:
            key_to_remove = list(self.td_coordinate_values.keys())[selected_index[0]]
            del self.td_coordinate_values[key_to_remove]
            self.update_td_listbox()
            self.update_top_tag()

    def remove_from_pc(self):
        selected_index = self.pc_listbox.curselection()
        if selected_index:
            key_to_remove = list(self.pc_coordinate_values.keys())[selected_index[0]]
            del self.pc_coordinate_values[key_to_remove]
            self.update_pc_listbox()

    def clear_td(self):
        self.td_coordinate_values.clear()
        self.update_td_listbox()
        self.update_top_tag()

    def clear_pc(self):
        self.pc_coordinate_values.clear()
        self.update_pc_listbox()

    def update_top_tag(self):
        try:
            first_entry = self.td_listbox.get(0)
            self.top_tag = first_entry.split(':')[0]
        except:
            pass