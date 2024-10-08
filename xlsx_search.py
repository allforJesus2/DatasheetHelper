import os
import openpyxl
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText
import fitz
import docx
def search_doc(folder_path, keyword, search_str, comparison_method, file_types=['docx', 'pdf', 'xlsx'],
               traversal_method='listdir', exclude_word=''):
    if not os.path.exists(folder_path):
        messagebox.showerror("Error", f"Folder '{folder_path}' does not exist.")
        return []

    results = []

    if traversal_method == 'listdir':
        for filename in os.listdir(folder_path):
            if not filename.startswith('~') and keyword in filename and (not exclude_word or exclude_word not in filename):
                if 'xlsx' in file_types and filename.endswith(".xlsx") or filename.endswith(".xlsm"):
                    print('checking ', filename)
                    file_path = os.path.join(folder_path, filename)
                    process_xlsx(file_path, search_str, comparison_method, results)
                elif 'docx' in file_types and filename.endswith(".docx"):
                    print('checking ', filename)
                    file_path = os.path.join(folder_path, filename)
                    process_docx(file_path, search_str, comparison_method, results)
                elif 'pdf' in file_types and filename.endswith(".pdf"):
                    print('checking ', filename)
                    file_path = os.path.join(folder_path, filename)
                    print(file_path)
                    process_pdf(file_path, search_str, comparison_method, results)
    elif traversal_method == 'os_walk':
        for root, dirs, files in os.walk(folder_path):
            for filename in files:
                if not filename.startswith('~') and keyword in filename and (not exclude_word or exclude_word not in filename):
                    if 'xlsx' in file_types and filename.endswith(".xlsx") or filename.endswith(".xlsm"):
                        print('checking ', filename)
                        file_path = os.path.join(root, filename)
                        process_xlsx(file_path, search_str, comparison_method, results)
                    elif 'docx' in file_types and filename.endswith(".docx"):
                        print('checking ', filename)
                        file_path = os.path.join(root, filename)
                        process_docx(file_path, search_str, comparison_method, results)
                    elif 'pdf' in file_types and filename.endswith(".pdf"):
                        print('checking ', filename)
                        file_path = os.path.join(root, filename)
                        process_pdf(file_path, search_str, comparison_method, results)

    else:
        messagebox.showerror("Error", "Invalid traversal method specified.")

    return results


def process_xlsx(file_path, cell_value, comparison_method, results):
    try:
        workbook = openpyxl.load_workbook(file_path, data_only=True, read_only=True)
        for sheet in workbook:
            for row in sheet.rows:
                for cell in row:
                    cell_str = str(cell.value)
                    if comparison_method == 'Equals' and cell_value == cell_str:
                        results.append(f"Found '{cell_value}' in '{file_path}' on sheet {sheet.title}: {cell.coordinate}\n")
                    elif comparison_method == 'Contains' and cell_value in cell_str:
                        results.append(f"Found '{cell_value}' in {cell_str} in {file_path} on sheet {sheet.title}: {cell.coordinate}\n")
        workbook.close()
    except Exception as e:
        results.append(f"Error processing '{file_path}': {e}")


def process_pdf(file_path, cell_value, comparison_method, results):
    try:
        with fitz.open(file_path) as pdf:
            for page in pdf:
                text = page.get_text()
                if comparison_method == 'Equals':
                    if f" {cell_value} " in text:  # search for exact word within delimiters
                        results.append(f"Found '{cell_value}' in '{file_path}' on page {page+1}\n")
                elif comparison_method == 'Contains':
                    if cell_value in text:  # search for string sequence
                        results.append(f"Found '{cell_value}' in '{file_path}' on page {page+1}\n")

    except Exception as e:
        results.append(f"Error processing '{file_path}': {e}")




def process_docx(file_path, cell_value, comparison_method, results):
    try:
        doc = docx.Document(file_path)
        for para in doc.paragraphs:
            text = para.text
            if comparison_method == 'Equals':
                if f" {cell_value} " in text:  # search for exact word within delimiters
                    results.append(f"Found '{cell_value}' in '{file_path}'\n")
            elif comparison_method == 'Contains':
                if cell_value in text:  # search for string sequence
                    results.append(f"Found '{cell_value}' in '{file_path}'\n")
    except Exception as e:
        results.append(f"Error processing '{file_path}': {e}")

class ExcelSearchApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Excel Cell Search")
        self.geometry("800x600")
        self.configure(padx=20, pady=20)
        self.create_widgets()

    def create_widgets(self):
        # Input Frame
        input_frame = ttk.LabelFrame(self, text="Search Parameters", padding=(10, 5))
        input_frame.pack(fill=tk.X, padx=10, pady=10)

        ttk.Label(input_frame, text="Filename Keyword:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.keyword_entry = ttk.Entry(input_frame, width=30)
        self.keyword_entry.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)

        ttk.Label(input_frame, text="Exclusion Keyword:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.exclude_entry = ttk.Entry(input_frame, width=30)
        self.exclude_entry.grid(row=1, column=1, sticky=tk.W, padx=5, pady=5)

        ttk.Label(input_frame, text="Cell Value to Search:").grid(row=2, column=0, sticky=tk.W, padx=5, pady=5)
        self.cell_value_entry = ttk.Entry(input_frame, width=30)
        self.cell_value_entry.grid(row=2, column=1, sticky=tk.W, padx=5, pady=5)

        ttk.Label(input_frame, text="Comparison Method:").grid(row=3, column=0, sticky=tk.W, padx=5, pady=5)
        self.comparison_method = ttk.Combobox(input_frame, values=['Equals', 'Contains'], width=27)
        self.comparison_method.current(0)
        self.comparison_method.grid(row=3, column=1, sticky=tk.W, padx=5, pady=5)

        ttk.Label(input_frame, text="Traversal Method:").grid(row=4, column=0, sticky=tk.W, padx=5, pady=5)
        self.traversal_method = ttk.Combobox(input_frame, values=['listdir', 'os_walk'], width=27)
        self.traversal_method.current(0)
        self.traversal_method.grid(row=4, column=1, sticky=tk.W, padx=5, pady=5)

        # File Type Selection Frame
        file_type_frame = ttk.LabelFrame(self, text="File Types", padding=(10, 5))
        file_type_frame.pack(fill=tk.X, padx=10, pady=10)

        self.xlsx_var = tk.IntVar()
        ttk.Checkbutton(file_type_frame, text="xlsx", variable=self.xlsx_var).grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)

        self.docx_var = tk.IntVar()
        ttk.Checkbutton(file_type_frame, text="docx", variable=self.docx_var).grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)

        self.pdf_var = tk.IntVar()
        ttk.Checkbutton(file_type_frame, text="pdf", variable=self.pdf_var).grid(row=0, column=2, sticky=tk.W, padx=5, pady=5)

        # Folder Selection Frame
        folder_frame = ttk.LabelFrame(self, text="Folder Selection", padding=(10, 5))
        folder_frame.pack(fill=tk.X, padx=10, pady=10)

        ttk.Label(folder_frame, text="Single Folder:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.folder_path_entry = ttk.Entry(folder_frame, width=50)
        self.folder_path_entry.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
        ttk.Button(folder_frame, text="Browse", command=self.browse_folder).grid(row=0, column=2, padx=5, pady=5)
        ttk.Button(folder_frame, text="Search", command=self.search_button_click).grid(row=0, column=3, padx=5, pady=5)

        ttk.Label(folder_frame, text="Multiple Folders:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.focused_folder_paths = ScrolledText(folder_frame, width=50, height=3)
        self.focused_folder_paths.grid(row=1, column=1, columnspan=2, sticky=tk.W, padx=5, pady=5)
        ttk.Button(folder_frame, text="Focused Search", command=self.focused_search_button_click).grid(row=1, column=3, padx=5, pady=5)

        # Results Frame
        results_frame = ttk.LabelFrame(self, text="Search Results", padding=(10, 5))
        results_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.result_text = ScrolledText(results_frame, wrap=tk.WORD, width=70, height=10)
        self.result_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.result_text.config()

    def browse_folder(self):
        folder_path = filedialog.askdirectory()
        self.folder_path_entry.delete(0, tk.END)
        self.folder_path_entry.insert(0, folder_path)

    def search_button_click(self):
        folder_path = self.folder_path_entry.get()
        self.perform_search(folder_path)

    def focused_search_button_click(self):
        folder_paths = self.focused_folder_paths.get('1.0', tk.END).strip().split('\n')
        for folder_path in folder_paths:
            if folder_path.strip():
                self.perform_search(folder_path.strip())

    def perform_search(self, folder_path):
        keyword = self.keyword_entry.get()
        exclude_word = self.exclude_entry.get()
        cell_value = self.cell_value_entry.get()
        comparison_method = self.comparison_method.get()
        traversal_method = self.traversal_method.get()
        file_types = []
        if self.xlsx_var.get():
            file_types.append('xlsx')
        if self.docx_var.get():
            file_types.append('docx')
        if self.pdf_var.get():
            file_types.append('pdf')

        results = search_doc(folder_path, keyword, cell_value, comparison_method, file_types=file_types, traversal_method=traversal_method, exclude_word=exclude_word)

        self.result_text.config(state='normal')
        #self.result_text.insert(tk.END, f"Search results for folder: {folder_path}\n")
        if results:
            for result in results:
                self.result_text.insert(tk.END, result + '\n')
        else:
            pass
            #self.result_text.insert(tk.END, "No matching cells found.\n")
        self.result_text.insert(tk.END, "\n")
        self.result_text.config()
        self.result_text.see(tk.END)

if __name__ == "__main__":
    app = ExcelSearchApp()
    app.mainloop()