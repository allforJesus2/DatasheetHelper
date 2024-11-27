import tkinter as tk
from tkinter import ttk
from typing import Dict, List, Tuple, Callable

class XKeySelector:
    def __init__(self, parent: tk.Tk, td_data: Dict, current_xkey: str, transformation_code: str,
                 on_save: Callable[[str, str], None]):
        self.window = tk.Toplevel(parent)
        self.window.title("Configure translation_lambda")

        frame1 = tk.Frame(self.window)
        frame1.pack(fill="x", pady=5, padx=5)
        frame2 = tk.Frame(self.window, pady=5, padx=5)
        frame2.pack(fill="x")

        # Index Source Key selection
        tk.Label(frame1, text="Index Source Key: ").pack(side='left', padx=5)
        td_combo_values = []
        for key, value in td_data.items():
            td_combo_values = list(value.keys())
            break

        self.td_combo = ttk.Combobox(frame1, values=td_combo_values)
        self.td_combo.insert(0, current_xkey)
        self.td_combo.pack(side='left', fill="x", expand=True)

        # Transformation code entry
        tk.Label(frame2, text="PC key, x = Index Value: ").pack(side='left', padx=5)
        self.transformation_entry = tk.Entry(frame2)
        self.transformation_entry.insert(0, transformation_code)
        self.transformation_entry.pack(side='left', fill="x", expand=True)

        tk.Button(self.window, text="Save",
                  command=lambda: self._save_and_close(on_save)).pack(expand=True)

    def _save_and_close(self, on_save: Callable[[str, str], None]):
        on_save(self.td_combo.get(), self.transformation_entry.get())
        self.window.destroy()


class DatasheetConfigurator:
    def __init__(self, parent: tk.Tk, current_config: Dict, sheet_names: List[str],
                 on_save: Callable[[Dict], None]):
        self.window = tk.Toplevel()
        self.window.title("Configure Add Datasheet")

        fields = [
            ("Source Sheet Name:", "source_sheet_name", ttk.Combobox, {"values": sheet_names}),
            ("Datasheet Coordinates:", "datasheet_coord", tk.Entry, {}),
            ("Datasheet String:", "ds_str", tk.Entry, {}),
            ("Tag Pattern (Regex):", "tag_pattern", tk.Entry, {}),
            ("Top Tag Coordinate:", "top_tag", tk.Entry, {}),
            ("Rows per Sheet:", "rows_per_sheet", tk.Entry, {})
        ]

        self.entries = {}
        for label, key, widget_class, widget_kwargs in fields:
            tk.Label(self.window, text=label).pack(anchor='sw', pady=(10, 2))
            self.entries[key] = widget_class(self.window, **widget_kwargs)
            self.entries[key].insert(tk.END, str(current_config.get(key, "")))
            self.entries[key].pack(fill='x')

        tk.Button(self.window, text="Save",
                  command=lambda: self._save_and_close(on_save)).pack(pady=(10, 2))

    def _save_and_close(self, on_save: Callable[[Dict], None]):
        config = {
            "source_sheet_name": self.entries["source_sheet_name"].get(),
            "datasheet_coord": self.entries["datasheet_coord"].get(),
            "ds_str": self.entries["ds_str"].get(),
            "tag_pattern": self.entries["tag_pattern"].get(),
            "top_tag": self.entries["top_tag"].get(),
            "rows_per_sheet": int(self.entries["rows_per_sheet"].get())
        }
        on_save(config)
        self.window.destroy()


class TagFilterConfigurator:
    def __init__(self, parent: tk.Tk, td_data: Dict, current_filters: List[List[str]],
                 on_save: Callable[[List[List[str]]], None]):
        self.window = tk.Toplevel(parent)
        self.window.title("Set Tag Filters")

        self.td_keys = []
        for key, value in td_data.items():
            self.td_keys = list(value.keys())
            break

        self.filters_entries = []
        self.on_save = on_save

        control_frame = tk.Frame(self.window)
        control_frame.grid(row=0, column=0, columnspan=4, sticky='ew', padx=5, pady=5)

        tk.Button(control_frame, text="Add New",
                  command=self._add_filter_row).pack(side=tk.LEFT, expand=True, fill='x', padx=5)
        tk.Button(control_frame, text="Save",
                  command=self._save_filters).pack(side=tk.LEFT, expand=True, fill='x', padx=5)

        for name, filter_value in current_filters:
            self._add_filter_row(name, filter_value)
        self._add_filter_row()

    def _add_filter_row(self, name='', filter_value=''):
        row = len(self.filters_entries) + 1

        name_label = tk.Label(self.window, text=f"Index Key {row}:")
        name_label.grid(row=row, column=0)
        name_entry = ttk.Combobox(self.window, values=self.td_keys)
        name_entry.grid(row=row, column=1)
        name_entry.set(name)

        filter_label = tk.Label(self.window, text=f"Filter {row}:")
        filter_label.grid(row=row, column=2)
        filter_entry = tk.Entry(self.window)
        filter_entry.grid(row=row, column=3)
        filter_entry.insert(0, filter_value)

        self.filters_entries.append((name_entry, filter_entry))

    def _save_filters(self):
        filters = []
        for name_entry, filter_entry in self.filters_entries:
            name = name_entry.get()
            filter_value = filter_entry.get()
            if name and filter_value:
                filters.append([name, filter_value])
        self.on_save(filters)
        self.window.destroy()