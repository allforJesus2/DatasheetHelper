import os
import openpyxl
import xlwings as xw
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import json

class ExcelEditorApp:
    def __init__(self, master):
        self.master = master
        self.master.title("Excel Editor")
        self.master.geometry("900x500")

        self.edits = []
        self.create_widgets()

    def create_widgets(self):
        # File selection frame
        file_frame = ttk.LabelFrame(self.master, text="File Selection")
        file_frame.pack(padx=10, pady=10, fill=tk.X)

        self.file_selection = tk.StringVar(value="WALK")
        ttk.Radiobutton(file_frame, text="Walk Directory", variable=self.file_selection, value="WALK").pack(
            side=tk.LEFT)
        ttk.Radiobutton(file_frame, text="Single Folder", variable=self.file_selection, value="FOLDER").pack(
            side=tk.LEFT)
        ttk.Radiobutton(file_frame, text="Single File", variable=self.file_selection, value="FILE").pack(side=tk.LEFT)

        # Edits frame
        edits_frame = ttk.LabelFrame(self.master, text="Edits")
        edits_frame.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

        # Edits table
        self.edits_tree = ttk.Treeview(edits_frame, columns=("Keyword", "Col Offset", "Row Offset", "New Value"),
                                       show="headings")
        self.edits_tree.heading("Keyword", text="Keyword")
        self.edits_tree.heading("Col Offset", text="Col Offset")
        self.edits_tree.heading("Row Offset", text="Row Offset")
        self.edits_tree.heading("New Value", text="New Value")
        self.edits_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(edits_frame, orient=tk.VERTICAL, command=self.edits_tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.edits_tree.configure(yscrollcommand=scrollbar.set)

        # Edit input frame
        input_frame = ttk.Frame(self.master)
        input_frame.pack(padx=10, pady=10, fill=tk.X)

        ttk.Label(input_frame, text="Keyword:").grid(row=0, column=0)
        self.keyword_entry = ttk.Entry(input_frame)
        self.keyword_entry.grid(row=0, column=1)

        ttk.Label(input_frame, text="Col Offset:").grid(row=0, column=2)
        self.col_offset_entry = ttk.Entry(input_frame, width=5)
        self.col_offset_entry.grid(row=0, column=3)

        ttk.Label(input_frame, text="Row Offset:").grid(row=0, column=4)
        self.row_offset_entry = ttk.Entry(input_frame, width=5)
        self.row_offset_entry.grid(row=0, column=5)

        ttk.Label(input_frame, text="New Value:").grid(row=0, column=6)
        self.new_value_entry = ttk.Entry(input_frame)
        self.new_value_entry.grid(row=0, column=7)

        ttk.Button(input_frame, text="Add Edit", command=self.add_edit).grid(row=0, column=8, padx=5)

        # Action buttons
        action_frame = ttk.Frame(self.master)
        action_frame.pack(pady=10)

        ttk.Button(action_frame, text="Process Files", command=self.process_files).pack(side=tk.LEFT, padx=5)
        ttk.Button(action_frame, text="Clear Edits", command=self.clear_edits).pack(side=tk.LEFT, padx=5)

        # Add Import/Export buttons
        import_export_frame = ttk.Frame(self.master)
        import_export_frame.pack(pady=5)

        ttk.Button(import_export_frame, text="Import Edits", command=self.import_edits).pack(side=tk.LEFT, padx=5)
        ttk.Button(import_export_frame, text="Export Edits", command=self.export_edits).pack(side=tk.LEFT, padx=5)

    def import_edits(self):
        file_path = filedialog.askopenfilename(title="Import Edits", filetypes=[("JSON files", "*.json")])
        if file_path:
            try:
                with open(file_path, 'r') as f:
                    imported_edits = json.load(f)
                self.edits = imported_edits
                self.refresh_edits_tree()
                messagebox.showinfo("Import Successful", f"Imported {len(self.edits)} edits from {file_path}")
            except Exception as e:
                messagebox.showerror("Import Error", f"Failed to import edits: {str(e)}")

    def export_edits(self):
        if not self.edits:
            messagebox.showwarning("No Edits", "There are no edits to export.")
            return

        file_path = filedialog.asksaveasfilename(title="Export Edits", defaultextension=".json", filetypes=[("JSON files", "*.json")])
        if file_path:
            try:
                with open(file_path, 'w') as f:
                    json.dump(self.edits, f, indent=2)
                messagebox.showinfo("Export Successful", f"Exported {len(self.edits)} edits to {file_path}")
            except Exception as e:
                messagebox.showerror("Export Error", f"Failed to export edits: {str(e)}")

    def refresh_edits_tree(self):
        self.edits_tree.delete(*self.edits_tree.get_children())
        for edit in self.edits:
            self.edits_tree.insert("", tk.END, values=edit)

    def add_edit(self):
        keyword = self.keyword_entry.get()
        col_offset = self.col_offset_entry.get()
        row_offset = self.row_offset_entry.get()
        new_value = self.new_value_entry.get()

        if keyword and col_offset and row_offset and new_value:
            self.edits.append((keyword, int(col_offset), int(row_offset), new_value))
            self.edits_tree.insert("", tk.END, values=(keyword, col_offset, row_offset, new_value))
            self.clear_input_fields()
        else:
            messagebox.showwarning("Invalid Input", "Please fill all fields.")

    def clear_input_fields(self):
        self.keyword_entry.delete(0, tk.END)
        self.col_offset_entry.delete(0, tk.END)
        self.row_offset_entry.delete(0, tk.END)
        self.new_value_entry.delete(0, tk.END)

    def clear_edits(self):
        self.edits = []
        self.edits_tree.delete(*self.edits_tree.get_children())

    def process_files(self):
        selection = self.file_selection.get()
        if selection == "WALK":
            root_folder = filedialog.askdirectory(title='Select folder to walk through')
            self.process_walk(root_folder)
        elif selection == "FOLDER":
            folder = filedialog.askdirectory(title='Select folder with Excel files')
            self.process_folder(folder)
        elif selection == "FILE":
            file_path = filedialog.askopenfilename(title="Select Excel file", filetypes=[("Excel files", "*.xlsx")])
            self.process_file(file_path)

    def process_walk(self, root_folder):
        for root, _, files in os.walk(root_folder):
            for file in files:
                if file.endswith('.xlsx') and not file.startswith('~'):
                    file_path = os.path.join(root, file)
                    self.make_edits(file_path)

    def process_folder(self, folder):
        for file in os.listdir(folder):
            if file.endswith('.xlsx') and not file.startswith('~'):
                file_path = os.path.join(folder, file)
                self.make_edits(file_path)

    def process_file(self, file_path):
        self.make_edits(file_path)

    def make_edits(self, file_path):
        try:
            app = xw.App(visible=False)
            wb = app.books.open(file_path)
            for edit in self.edits:
                keyword, col_offset, row_offset, new_value = edit
                self.edit_excel_cell(wb, keyword, col_offset, row_offset, new_value)
            wb.save()
            wb.close()
            app.quit()
            print(f"Processed: {file_path}")
        except Exception as e:
            print(f"Error processing {file_path}: {str(e)}")

    def edit_excel_cell(self, wb, keyword, col_offset, row_offset, new_value):
        for sheet in wb.sheets:
            used_range = sheet.used_range
            values = used_range.value
            for i, row in enumerate(values):
                for j, cell_value in enumerate(row):
                    if cell_value == keyword:
                        target_cell = sheet.cells(i + 1 + row_offset, j + 2 + col_offset)
                        target_cell.value = new_value
                        print(f"Updated cell {target_cell.address} in sheet {sheet.name}")


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelEditorApp(root)
    root.mainloop()