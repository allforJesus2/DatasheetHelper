import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import re
import os
from pathlib import Path
import threading
import queue

# Excel handling
try:
    import openpyxl
    from openpyxl import load_workbook
    EXCEL_AVAILABLE = True
except ImportError:
    EXCEL_AVAILABLE = False

# Word handling
try:
    from docx import Document
    WORD_AVAILABLE = True
except ImportError:
    WORD_AVAILABLE = False

class DocumentSearchReplace:
    def __init__(self, root):
        self.root = root
        self.root.title("Document Search & Replace Tool")
        self.root.geometry("800x600")
        self.root.configure(bg='#f0f0f0')
        
        # Configure style
        self.style = ttk.Style()
        self.style.theme_use('clam')
        
        # Variables
        self.selected_file = tk.StringVar()
        self.search_text = tk.StringVar()
        self.replace_text = tk.StringVar()
        self.use_regex = tk.BooleanVar()
        self.case_sensitive = tk.BooleanVar()
        self.search_results = []
        self.current_file_type = None
        
        # Results queue for threading
        self.results_queue = queue.Queue()
        
        self.create_widgets()
        self.check_dependencies()
    
    def check_dependencies(self):
        """Check if required libraries are available"""
        missing_libs = []
        if not EXCEL_AVAILABLE:
            missing_libs.append("openpyxl (for Excel files)")
        if not WORD_AVAILABLE:
            missing_libs.append("python-docx (for Word files)")
        
        if missing_libs:
            messagebox.showwarning(
                "Missing Dependencies",
                f"The following libraries are not installed:\n{', '.join(missing_libs)}\n\n"
                "Please install them using:\npip install openpyxl python-docx"
            )
    
    def create_widgets(self):
        """Create the main UI widgets"""
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # File selection section
        self.create_file_section(main_frame)
        
        # Search options section
        self.create_search_section(main_frame)
        
        # Results section
        self.create_results_section(main_frame)
        
        # Buttons section
        self.create_buttons_section(main_frame)
    
    def create_file_section(self, parent):
        """Create file selection widgets"""
        # File selection frame
        file_frame = ttk.LabelFrame(parent, text="Document Selection", padding="10")
        file_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        file_frame.columnconfigure(1, weight=1)
        
        # File path label and entry
        ttk.Label(file_frame, text="Document:").grid(row=0, column=0, sticky=tk.W, padx=(0, 5))
        file_entry = ttk.Entry(file_frame, textvariable=self.selected_file, width=50)
        file_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 5))
        
        # Browse button
        browse_btn = ttk.Button(file_frame, text="Browse", command=self.browse_file)
        browse_btn.grid(row=0, column=2)
        
        # File info label
        self.file_info_label = ttk.Label(file_frame, text="No file selected", foreground="gray")
        self.file_info_label.grid(row=1, column=0, columnspan=3, sticky=tk.W, pady=(5, 0))
    
    def create_search_section(self, parent):
        """Create search and replace widgets"""
        # Search frame
        search_frame = ttk.LabelFrame(parent, text="Search & Replace", padding="10")
        search_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        search_frame.columnconfigure(1, weight=1)
        
        # Search text
        ttk.Label(search_frame, text="Search for:").grid(row=0, column=0, sticky=tk.W, padx=(0, 5))
        search_entry = ttk.Entry(search_frame, textvariable=self.search_text, width=50)
        search_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 5))
        
        # Replace text
        ttk.Label(search_frame, text="Replace with:").grid(row=1, column=0, sticky=tk.W, padx=(0, 5), pady=(10, 0))
        replace_entry = ttk.Entry(search_frame, textvariable=self.replace_text, width=50)
        replace_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), padx=(0, 5), pady=(10, 0))
        
        # Options frame
        options_frame = ttk.Frame(search_frame)
        options_frame.grid(row=2, column=0, columnspan=2, sticky=tk.W, pady=(10, 0))
        
        # Checkboxes
        regex_cb = ttk.Checkbutton(options_frame, text="Use Regular Expression", variable=self.use_regex)
        regex_cb.grid(row=0, column=0, sticky=tk.W, padx=(0, 20))
        
        case_cb = ttk.Checkbutton(options_frame, text="Case Sensitive", variable=self.case_sensitive)
        case_cb.grid(row=0, column=1, sticky=tk.W)
    
    def create_results_section(self, parent):
        """Create results display widgets"""
        # Results frame
        results_frame = ttk.LabelFrame(parent, text="Search Results", padding="10")
        results_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        results_frame.columnconfigure(0, weight=1)
        results_frame.rowconfigure(0, weight=1)
        parent.rowconfigure(2, weight=1)
        
        # Results text area
        self.results_text = scrolledtext.ScrolledText(
            results_frame, 
            height=15, 
            width=80,
            wrap=tk.WORD,
            font=('Consolas', 9)
        )
        self.results_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Status label
        self.status_label = ttk.Label(results_frame, text="Ready", foreground="gray")
        self.status_label.grid(row=1, column=0, sticky=tk.W, pady=(5, 0))
    
    def create_buttons_section(self, parent):
        """Create action buttons"""
        # Buttons frame
        buttons_frame = ttk.Frame(parent)
        buttons_frame.grid(row=3, column=0, columnspan=2, pady=(0, 10))
        
        # Search button
        self.search_btn = ttk.Button(buttons_frame, text="Search", command=self.search_document)
        self.search_btn.grid(row=0, column=0, padx=(0, 10))
        
        # Replace button
        self.replace_btn = ttk.Button(buttons_frame, text="Replace All", command=self.replace_all)
        self.replace_btn.grid(row=0, column=1, padx=(0, 10))
        
        # Clear button
        clear_btn = ttk.Button(buttons_frame, text="Clear Results", command=self.clear_results)
        clear_btn.grid(row=0, column=2)
    
    def browse_file(self):
        """Open file dialog to select document"""
        file_types = []
        if EXCEL_AVAILABLE:
            file_types.append(("Excel files", "*.xlsx *.xls"))
        if WORD_AVAILABLE:
            file_types.append(("Word files", "*.docx *.doc"))
        
        if not file_types:
            messagebox.showerror("Error", "No supported file types available. Please install required libraries.")
            return
        
        file_types.append(("All files", "*.*"))
        
        filename = filedialog.askopenfilename(
            title="Select Document",
            filetypes=file_types
        )
        
        if filename:
            self.selected_file.set(filename)
            self.update_file_info(filename)
    
    def update_file_info(self, filename):
        """Update file information display"""
        try:
            file_path = Path(filename)
            if file_path.exists():
                size = file_path.stat().st_size
                size_str = f"{size:,} bytes" if size < 1024 else f"{size/1024:.1f} KB"
                
                # Determine file type
                if filename.lower().endswith(('.xlsx', '.xls')):
                    self.current_file_type = 'excel'
                    file_type = "Excel Document"
                elif filename.lower().endswith(('.docx', '.doc')):
                    self.current_file_type = 'word'
                    file_type = "Word Document"
                else:
                    self.current_file_type = None
                    file_type = "Unknown"
                
                self.file_info_label.config(
                    text=f"Type: {file_type} | Size: {size_str} | Path: {file_path.name}",
                    foreground="black"
                )
            else:
                self.file_info_label.config(text="File not found", foreground="red")
        except Exception as e:
            self.file_info_label.config(text=f"Error reading file: {str(e)}", foreground="red")
    
    def search_document(self):
        """Search for text in the selected document"""
        if not self.selected_file.get():
            messagebox.showwarning("Warning", "Please select a document first.")
            return
        
        if not self.search_text.get():
            messagebox.showwarning("Warning", "Please enter search text.")
            return
        
        # Clear previous results
        self.clear_results()
        self.status_label.config(text="Searching...", foreground="blue")
        
        # Run search in separate thread
        search_thread = threading.Thread(target=self._search_worker)
        search_thread.daemon = True
        search_thread.start()
        
        # Start checking for results
        self.root.after(100, self.check_search_results)
    
    def _search_worker(self):
        """Worker function to perform search in background"""
        try:
            if self.current_file_type == 'excel':
                results = self.search_excel()
            elif self.current_file_type == 'word':
                results = self.search_word()
            else:
                results = []
            
            self.results_queue.put(('results', results))
        except Exception as e:
            self.results_queue.put(('error', str(e)))
    
    def check_search_results(self):
        """Check for search results from worker thread"""
        try:
            result_type, data = self.results_queue.get_nowait()
            
            if result_type == 'results':
                self.search_results = data
                self.display_results()
                count = len(data)
                self.status_label.config(
                    text=f"Found {count} match{'es' if count != 1 else ''}",
                    foreground="green" if count > 0 else "orange"
                )
            elif result_type == 'error':
                self.status_label.config(text=f"Error: {data}", foreground="red")
                messagebox.showerror("Search Error", f"An error occurred during search:\n{data}")
        
        except queue.Empty:
            # Continue checking
            self.root.after(100, self.check_search_results)
    
    def search_excel(self):
        """Search in Excel document"""
        if not EXCEL_AVAILABLE:
            raise Exception("openpyxl library not available")
        
        results = []
        workbook = load_workbook(self.selected_file.get(), data_only=True)
        
        for sheet_name in workbook.sheetnames:
            sheet = workbook[sheet_name]
            for row_idx, row in enumerate(sheet.iter_rows(values_only=True), 1):
                for col_idx, cell_value in enumerate(row, 1):
                    if cell_value is not None:
                        cell_str = str(cell_value)
                        matches = self.find_matches(cell_str)
                        for match in matches:
                            results.append({
                                'type': 'excel',
                                'sheet': sheet_name,
                                'row': row_idx,
                                'col': col_idx,
                                'cell_value': cell_str,
                                'match': match
                            })
        
        return results
    
    def search_word(self):
        """Search in Word document"""
        if not WORD_AVAILABLE:
            raise Exception("python-docx library not available")
        
        results = []
        doc = Document(self.selected_file.get())
        
        for para_idx, paragraph in enumerate(doc.paragraphs, 1):
            if paragraph.text:
                matches = self.find_matches(paragraph.text)
                for match in matches:
                    results.append({
                        'type': 'word',
                        'paragraph': para_idx,
                        'text': paragraph.text,
                        'match': match
                    })
        
        # Search in tables
        for table_idx, table in enumerate(doc.tables, 1):
            for row_idx, row in enumerate(table.rows, 1):
                for col_idx, cell in enumerate(row.cells, 1):
                    if cell.text:
                        matches = self.find_matches(cell.text)
                        for match in matches:
                            results.append({
                                'type': 'word_table',
                                'table': table_idx,
                                'row': row_idx,
                                'col': col_idx,
                                'text': cell.text,
                                'match': match
                            })
        
        return results
    
    def find_matches(self, text):
        """Find all matches in text based on search options"""
        search_pattern = self.search_text.get()
        
        if self.use_regex.get():
            try:
                flags = 0 if self.case_sensitive.get() else re.IGNORECASE
                pattern = re.compile(search_pattern, flags)
                matches = list(pattern.finditer(text))
                return [{'start': m.start(), 'end': m.end(), 'text': m.group()} for m in matches]
            except re.error as e:
                raise Exception(f"Invalid regular expression: {str(e)}")
        else:
            search_text = search_pattern
            if not self.case_sensitive.get():
                search_text = search_text.lower()
                text_lower = text.lower()
            else:
                text_lower = text
            
            matches = []
            start = 0
            while True:
                pos = text_lower.find(search_text, start)
                if pos == -1:
                    break
                matches.append({
                    'start': pos,
                    'end': pos + len(search_text),
                    'text': text[pos:pos + len(search_text)]
                })
                start = pos + 1
            
            return matches
    
    def display_results(self):
        """Display search results in the text area"""
        self.results_text.delete(1.0, tk.END)
        
        if not self.search_results:
            self.results_text.insert(tk.END, "No matches found.\n")
            return
        
        for i, result in enumerate(self.search_results, 1):
            if result['type'] == 'excel':
                self.results_text.insert(tk.END, 
                    f"{i}. Excel - Sheet: {result['sheet']}, Cell: {result['row']}:{result['col']}\n")
                self.results_text.insert(tk.END, f"   Value: {result['cell_value']}\n")
                self.results_text.insert(tk.END, f"   Match: {result['match']['text']}\n\n")
            
            elif result['type'] == 'word':
                self.results_text.insert(tk.END, 
                    f"{i}. Word - Paragraph {result['paragraph']}\n")
                self.results_text.insert(tk.END, f"   Text: {result['text']}\n")
                self.results_text.insert(tk.END, f"   Match: {result['match']['text']}\n\n")
            
            elif result['type'] == 'word_table':
                self.results_text.insert(tk.END, 
                    f"{i}. Word Table - Table {result['table']}, Cell: {result['row']}:{result['col']}\n")
                self.results_text.insert(tk.END, f"   Text: {result['text']}\n")
                self.results_text.insert(tk.END, f"   Match: {result['match']['text']}\n\n")
    
    def replace_all(self):
        """Replace all occurrences in the document"""
        if not self.search_results:
            messagebox.showwarning("Warning", "No search results to replace. Please search first.")
            return
        
        if not self.replace_text.get():
            messagebox.showwarning("Warning", "Please enter replacement text.")
            return
        
        # Confirm replacement
        count = len(self.search_results)
        response = messagebox.askyesno(
            "Confirm Replacement",
            f"Replace {count} occurrence{'s' if count != 1 else ''}?\n\n"
            "This action cannot be undone. Make sure to backup your document first."
        )
        
        if not response:
            return
        
        try:
            self.status_label.config(text="Replacing...", foreground="blue")
            
            if self.current_file_type == 'excel':
                self.replace_excel()
            elif self.current_file_type == 'word':
                self.replace_word()
            
            self.status_label.config(text="Replacement completed successfully!", foreground="green")
            messagebox.showinfo("Success", f"Successfully replaced {count} occurrence{'s' if count != 1 else ''}.")
            
            # Clear results after replacement
            self.search_results = []
            self.clear_results()
            
        except Exception as e:
            self.status_label.config(text=f"Error: {str(e)}", foreground="red")
            messagebox.showerror("Replacement Error", f"An error occurred during replacement:\n{str(e)}")
    
    def replace_excel(self):
        """Replace text in Excel document"""
        if not EXCEL_AVAILABLE:
            raise Exception("openpyxl library not available")
        
        workbook = load_workbook(self.selected_file.get())
        
        # Group results by sheet
        sheet_results = {}
        for result in self.search_results:
            if result['type'] == 'excel':
                sheet_name = result['sheet']
                if sheet_name not in sheet_results:
                    sheet_results[sheet_name] = []
                sheet_results[sheet_name].append(result)
        
        # Replace in each sheet
        for sheet_name, results in sheet_results.items():
            sheet = workbook[sheet_name]
            
            # Group by cell to avoid multiple replacements in same cell
            cell_results = {}
            for result in results:
                cell_key = (result['row'], result['col'])
                if cell_key not in cell_results:
                    cell_results[cell_key] = result['cell_value']
                
                # Replace in cell value
                cell_value = cell_results[cell_key]
                match = result['match']
                new_value = (
                    cell_value[:match['start']] + 
                    self.replace_text.get() + 
                    cell_value[match['end']:]
                )
                cell_results[cell_key] = new_value
            
            # Update cells
            for (row, col), new_value in cell_results.items():
                sheet.cell(row=row, column=col, value=new_value)
        
        # Save the workbook
        workbook.save(self.selected_file.get())
    
    def replace_word(self):
        """Replace text in Word document"""
        if not WORD_AVAILABLE:
            raise Exception("python-docx library not available")
        
        doc = Document(self.selected_file.get())
        
        # Replace in paragraphs
        for result in self.search_results:
            if result['type'] == 'word':
                para_idx = result['paragraph'] - 1
                if para_idx < len(doc.paragraphs):
                    paragraph = doc.paragraphs[para_idx]
                    match = result['match']
                    new_text = (
                        paragraph.text[:match['start']] + 
                        self.replace_text.get() + 
                        paragraph.text[match['end']:]
                    )
                    paragraph.text = new_text
        
        # Replace in table cells
        for result in self.search_results:
            if result['type'] == 'word_table':
                table_idx = result['table'] - 1
                row_idx = result['row'] - 1
                col_idx = result['col'] - 1
                
                if (table_idx < len(doc.tables) and 
                    row_idx < len(doc.tables[table_idx].rows) and
                    col_idx < len(doc.tables[table_idx].rows[row_idx].cells)):
                    
                    cell = doc.tables[table_idx].rows[row_idx].cells[col_idx]
                    match = result['match']
                    new_text = (
                        cell.text[:match['start']] + 
                        self.replace_text.get() + 
                        cell.text[match['end']:]
                    )
                    cell.text = new_text
        
        # Save the document
        doc.save(self.selected_file.get())
    
    def clear_results(self):
        """Clear the results display"""
        self.results_text.delete(1.0, tk.END)
        self.status_label.config(text="Ready", foreground="gray")

def main():
    """Main function to run the application"""
    root = tk.Tk()
    app = DocumentSearchReplace(root)
    
    # Center the window
    root.update_idletasks()
    x = (root.winfo_screenwidth() // 2) - (root.winfo_width() // 2)
    y = (root.winfo_screenheight() // 2) - (root.winfo_height() // 2)
    root.geometry(f"+{x}+{y}")
    
    root.mainloop()

if __name__ == "__main__":
    main()
