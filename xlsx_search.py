import os
import openpyxl
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText
import fitz
import docx
def search_doc(folder_path, keyword, search_str, comparison_method, file_types=['docx', 'pdf', 'xlsx'],
               traversal_method='listdir', exclude_word='', progress_callback=None, log_callback=None):
    def log(msg):
        if log_callback:
            log_callback(msg)
        else:
            print(msg)
    
    log(f"\n### SEARCH_DOC called ###")
    log(f"  Folder path: '{folder_path}'")
    log(f"  Folder exists: {os.path.exists(folder_path)}")
    
    if not os.path.exists(folder_path):
        messagebox.showerror("Error", f"Folder '{folder_path}' does not exist.")
        return []

    results = []
    
    # First, count total files to process
    files_to_process = []
    log(f"  Using traversal method: '{traversal_method}'")
    
    if traversal_method == 'listdir':
        log(f"  Listing files in folder...")
        all_files = os.listdir(folder_path)
        log(f"  Total files in folder: {len(all_files)}")
        for filename in all_files:
            log(f"    Checking: '{filename}'")
            if filename.startswith('~'):
                log(f"      -> Skipped (starts with ~)")
                continue
            if keyword and keyword not in filename:
                log(f"      -> Skipped (keyword '{keyword}' not in filename)")
                continue
            if exclude_word and exclude_word in filename:
                log(f"      -> Skipped (exclude word '{exclude_word}' found)")
                continue
            
            matched = False
            if 'xlsx' in file_types and (filename.endswith(".xlsx") or filename.endswith(".xlsm")):
                matched = True
                log(f"      -> Matched as XLSX file")
            elif 'docx' in file_types and filename.endswith(".docx"):
                matched = True
                log(f"      -> Matched as DOCX file")
            elif 'pdf' in file_types and filename.endswith(".pdf"):
                matched = True
                log(f"      -> Matched as PDF file")
            else:
                log(f"      -> Skipped (no file type match)")
            
            if matched:
                files_to_process.append((folder_path, filename))
    elif traversal_method == 'os_walk':
        for root, dirs, files in os.walk(folder_path):
            for filename in files:
                if not filename.startswith('~') and keyword in filename and (not exclude_word or exclude_word not in filename):
                    if ('xlsx' in file_types and (filename.endswith(".xlsx") or filename.endswith(".xlsm"))) or \
                       ('docx' in file_types and filename.endswith(".docx")) or \
                       ('pdf' in file_types and filename.endswith(".pdf")):
                        files_to_process.append((root, filename))
    
    total_files = len(files_to_process)
    log(f"\n  FILES TO PROCESS: {total_files}")
    if total_files > 0:
        log(f"  Files list:")
        for root, filename in files_to_process:
            log(f"    - {filename}")
    else:
        log(f"  WARNING: No files matched the criteria!")
    
    # Process files with progress updates
    log(f"\n  Starting file processing...")
    for idx, (root, filename) in enumerate(files_to_process):
        file_path = os.path.join(root, filename)
        log(f"    [{idx+1}/{total_files}] Processing: {filename}")
        
        if progress_callback:
            progress_callback(idx + 1, total_files, filename)
        
        if 'xlsx' in file_types and (filename.endswith(".xlsx") or filename.endswith(".xlsm")):
            log(f"      -> Processing as XLSX")
            process_xlsx(file_path, search_str, comparison_method, results)
        elif 'docx' in file_types and filename.endswith(".docx"):
            log(f"      -> Processing as DOCX")
            process_docx(file_path, search_str, comparison_method, results)
        elif 'pdf' in file_types and filename.endswith(".pdf"):
            log(f"      -> Processing as PDF")
            process_pdf(file_path, search_str, comparison_method, results)

    log(f"\n  SEARCH_DOC finished. Found {len(results)} results.")
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

class ProgressPopup(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Search Progress")
        self.geometry("400x120")
        self.resizable(False, False)
        
        # Center the window
        self.transient(parent)
        self.grab_set()
        
        # Progress label
        self.label = ttk.Label(self, text="Initializing search...", wraplength=380)
        self.label.pack(pady=10, padx=10)
        
        # Progress bar
        self.progress = ttk.Progressbar(self, length=350, mode='determinate')
        self.progress.pack(pady=10, padx=10)
        
        # Status label (current/total)
        self.status_label = ttk.Label(self, text="0 / 0 files")
        self.status_label.pack(pady=5)
        
    def update_progress(self, current, total, filename):
        self.progress['maximum'] = total
        self.progress['value'] = current
        self.label.config(text=f"Processing: {filename}")
        self.status_label.config(text=f"{current} / {total} files")
        self.update_idletasks()
    
    def close(self):
        self.grab_release()
        self.destroy()

class ExcelSearchApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Excel Cell Search")
        self.geometry("900x900")
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
        results_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        self.result_text = ScrolledText(results_frame, wrap=tk.WORD, width=70, height=8)
        self.result_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.result_text.config()
        
        # Log Frame
        log_frame = ttk.LabelFrame(self, text="Debug Log", padding=(10, 5))
        log_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        self.log_text = ScrolledText(log_frame, wrap=tk.WORD, width=70, height=8, bg='#f0f0f0', fg='#333333')
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.log_text.config(state='disabled')
        
        # Clear log button
        ttk.Button(log_frame, text="Clear Log", command=self.clear_log).pack(pady=5)

    def log(self, message):
        """Log message to both console and GUI log widget"""
        print(message)  # Still print to console
        self.log_text.config(state='normal')
        self.log_text.insert(tk.END, message + '\n')
        self.log_text.see(tk.END)
        self.log_text.config(state='disabled')
        self.update_idletasks()  # Force GUI update
    
    def clear_log(self):
        """Clear the debug log"""
        self.log_text.config(state='normal')
        self.log_text.delete('1.0', tk.END)
        self.log_text.config(state='disabled')
    
    def browse_folder(self):
        folder_path = filedialog.askdirectory()
        self.folder_path_entry.delete(0, tk.END)
        self.folder_path_entry.insert(0, folder_path)

    def search_button_click(self):
        folder_path = self.folder_path_entry.get()
        self.perform_search(folder_path)

    def focused_search_button_click(self):
        self.log("=" * 60)
        self.log("FOCUSED SEARCH CLICKED")
        self.log("=" * 60)
        raw_text = self.focused_folder_paths.get('1.0', tk.END)
        self.log(f"Raw text from text box: '{raw_text}'")
        folder_paths = raw_text.strip().replace('"', '').split('\n')
        self.log(f"Parsed folder paths: {folder_paths}")
        self.log(f"Number of folder paths: {len(folder_paths)}")
        
        for idx, folder_path in enumerate(folder_paths):
            self.log(f"\n--- Processing folder {idx + 1}/{len(folder_paths)} ---")
            self.log(f"Folder path (with whitespace): '{folder_path}'")
            if folder_path.strip():
                cleaned_path = folder_path.strip()
                self.log(f"Cleaned folder path: '{cleaned_path}'")
                self.perform_search(cleaned_path)
            else:
                self.log(f"Skipping empty folder path")
        self.log("\n" + "=" * 60)
        self.log("FOCUSED SEARCH COMPLETED")
        self.log("=" * 60)

    def perform_search(self, folder_path):
        self.log(f"\n>>> PERFORM_SEARCH called with folder: '{folder_path}'")
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

        self.log(f"  Keyword: '{keyword}'")
        self.log(f"  Exclude word: '{exclude_word}'")
        self.log(f"  Cell value to search: '{cell_value}'")
        self.log(f"  Comparison method: '{comparison_method}'")
        self.log(f"  Traversal method: '{traversal_method}'")
        self.log(f"  File types: {file_types}")

        # Create progress popup
        progress_popup = ProgressPopup(self)
        
        # Define progress callback
        def progress_callback(current, total, filename):
            progress_popup.update_progress(current, total, filename)
        
        try:
            self.log(f"\n  Calling search_doc...")
            results = search_doc(folder_path, keyword, cell_value, comparison_method, 
                               file_types=file_types, traversal_method=traversal_method, 
                               exclude_word=exclude_word, progress_callback=progress_callback,
                               log_callback=self.log)

            self.log(f"\n  Returned {len(results)} results from search_doc")
            
            self.result_text.config(state='normal')
            #self.result_text.insert(tk.END, f"Search results for folder: {folder_path}\n")
            if results:
                self.log(f"  Adding {len(results)} results to text widget")
                for result in results:
                    self.result_text.insert(tk.END, result + '\n')
            else:
                self.log(f"  No results to display")
                pass
                #self.result_text.insert(tk.END, "No matching cells found.\n")
            self.result_text.insert(tk.END, "\n")
            self.result_text.config()
            self.result_text.see(tk.END)
            self.log(f"  Results displayed in UI")
        finally:
            # Close progress popup
            self.log(f"  Closing progress popup")
            progress_popup.close()
            self.log(f"<<< PERFORM_SEARCH completed\n")

if __name__ == "__main__":
    app = ExcelSearchApp()
    app.mainloop()