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


class DataSourceRelationship:
    def __init__(self, source, target, key_transform):
        self.source = source  # Source DataSource object
        self.target = target  # Target DataSource object
        self.transform = key_transform  # Function to transform source key to target key


class DataSource:
    def __init__(self, name, file_path=None, headers=None):
        self.name = name
        self.file_path = file_path
        self.headers = headers or []
        self.data = {}

    def load_data(self):
        if self.file_path and self.headers:
            self.data = generate_dictionary_from_xlsx(self.file_path, self.headers)
        return self.data


class DataSourceManager:
    def __init__(self):
        self.sources = {}
        self.relationships = []

    def add_source(self, name, file_path=None, headers=None):
        source = DataSource(name, file_path, headers)
        self.sources[name] = source
        return source

    def add_relationship(self, source_name, target_name, transform_fn):
        if source_name in self.sources and target_name in self.sources:
            relationship = DataSourceRelationship(
                self.sources[source_name],
                self.sources[target_name],
                transform_fn
            )
            self.relationships.append(relationship)
            return relationship
        raise KeyError("Source or target not found")

    def get_related_data(self, source_name, source_key):
        results = {}
        for rel in self.relationships:
            if rel.source.name == source_name:
                try:
                    target_key = rel.transform(source_key)
                    if target_key in rel.target.data:
                        results[rel.target.name] = rel.target.data[target_key]
                except Exception as e:
                    print(f"Transform error: {e}")
        return results


class EnhancedDataGeneratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Enhanced Data Generator")
        self.data_manager = DataSourceManager()
        self.create_widgets()

    def create_widgets(self):
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(expand=True, fill='both', padx=5, pady=5)

        # Data Sources Tab
        sources_frame = ttk.Frame(self.notebook)
        self.notebook.add(sources_frame, text='Data Sources')

        # Data source controls
        ttk.Button(sources_frame, text="Add Source", command=self.add_source_dialog).pack(pady=5)
        self.sources_tree = ttk.Treeview(sources_frame, columns=('Path', 'Headers'), show='headings')
        self.sources_tree.heading('Path', text='File Path')
        self.sources_tree.heading('Headers', text='Headers')
        self.sources_tree.pack(expand=True, fill='both', pady=5)

        # Relationships Tab
        relations_frame = ttk.Frame(self.notebook)
        self.notebook.add(relations_frame, text='Relationships')

        ttk.Button(relations_frame, text="Add Relationship", command=self.add_relationship_dialog).pack(pady=5)
        self.relations_tree = ttk.Treeview(relations_frame, columns=('Source', 'Target', 'Transform'), show='headings')
        self.relations_tree.heading('Source', text='Source')
        self.relations_tree.heading('Target', text='Target')
        self.relations_tree.heading('Transform', text='Transform')
        self.relations_tree.pack(expand=True, fill='both', pady=5)

        # Output Tab
        output_frame = ttk.Frame(self.notebook)
        self.notebook.add(output_frame, text='Output')

        # Output configuration
        output_config = ttk.LabelFrame(output_frame, text="Output Configuration")
        output_config.pack(fill='x', padx=5, pady=5)

        ttk.Label(output_config, text="Template File:").pack(side='left', padx=5)
        self.template_entry = ttk.Entry(output_config)
        self.template_entry.pack(side='left', expand=True, fill='x', padx=5)
        ttk.Button(output_config, text="Browse", command=self.browse_template).pack(side='left', padx=5)

        # Generate button
        ttk.Button(output_frame, text="Generate Output", command=self.generate_output).pack(pady=10)

    def add_source_dialog(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("Add Data Source")

        ttk.Label(dialog, text="Name:").pack(pady=5)
        name_entry = ttk.Entry(dialog)
        name_entry.pack(pady=5)

        ttk.Label(dialog, text="File Path:").pack(pady=5)
        path_entry = ttk.Entry(dialog)
        path_entry.pack(pady=5)
        ttk.Button(dialog, text="Browse",
                   command=lambda: path_entry.insert(0, filedialog.askopenfilename())).pack()

        ttk.Label(dialog, text="Headers (comma-separated):").pack(pady=5)
        headers_entry = ttk.Entry(dialog)
        headers_entry.pack(pady=5)

        def save():
            name = name_entry.get()
            path = path_entry.get()
            headers = [h.strip() for h in headers_entry.get().split(',')]
            self.data_manager.add_source(name, path, headers)
            self.update_sources_tree()
            dialog.destroy()

        ttk.Button(dialog, text="Save", command=save).pack(pady=10)

    def add_relationship_dialog(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("Add Relationship")

        sources = list(self.data_manager.sources.keys())

        ttk.Label(dialog, text="Source:").pack(pady=5)
        source_combo = ttk.Combobox(dialog, values=sources)
        source_combo.pack(pady=5)

        ttk.Label(dialog, text="Target:").pack(pady=5)
        target_combo = ttk.Combobox(dialog, values=sources)
        target_combo.pack(pady=5)

        ttk.Label(dialog, text="Transform Function:").pack(pady=5)
        transform_entry = ttk.Entry(dialog)
        transform_entry.pack(pady=5)

        def save():
            source = source_combo.get()
            target = target_combo.get()
            transform_code = transform_entry.get()
            transform_fn = eval(f'lambda x: {transform_code}')
            self.data_manager.add_relationship(source, target, transform_fn)
            self.update_relations_tree()
            dialog.destroy()

        ttk.Button(dialog, text="Save", command=save).pack(pady=10)

    def update_sources_tree(self):
        self.sources_tree.delete(*self.sources_tree.get_children())
        for name, source in self.data_manager.sources.items():
            self.sources_tree.insert('', 'end', text=name,
                                     values=(source.file_path, ', '.join(source.headers)))

    def update_relations_tree(self):
        self.relations_tree.delete(*self.relations_tree.get_children())
        for rel in self.data_manager.relationships:
            self.relations_tree.insert('', 'end',
                                       values=(rel.source.name, rel.target.name, rel.transform.__code__))

    def browse_template(self):
        path = filedialog.askopenfilename()
        self.template_entry.delete(0, tk.END)
        self.template_entry.insert(0, path)

    def generate_output(self):
        template_path = self.template_entry.get()
        # Implement output generation logic here
        pass


if __name__ == "__main__":
    root = tk.Tk()
    app = EnhancedDataGeneratorApp(root)
    root.mainloop()