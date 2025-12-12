import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import sys
import win32com.client
import pythoncom
import pandas as pd
from PyPDF2 import PdfMerger, PdfReader, PdfWriter
import tempfile
import threading
from pathlib import Path
import json
from datetime import datetime
import pyperclip

class TabContent:
    """Class to hold the content of each tab"""
    def __init__(self, parent_frame):
        self.parent_frame = parent_frame
        self.selected_files = []
        self.setup_tab_ui()
        
    def setup_tab_ui(self):
        # Configure grid weights for the tab frame
        self.parent_frame.columnconfigure(0, weight=1)
        self.parent_frame.rowconfigure(2, weight=1)
        
        # Instructions
        instructions = ttk.Label(self.parent_frame, text="Select documents to merge. Files will appear in the order you add them.\nFirst file added = First page in final PDF", 
                               font=("Arial", 10), foreground="blue")
        instructions.grid(row=0, column=0, columnspan=3, pady=(0, 10))
        
        # Buttons frame
        button_frame = ttk.Frame(self.parent_frame)
        button_frame.grid(row=1, column=0, columnspan=3, pady=(0, 10), sticky=(tk.W, tk.E))
        
        # Add files button
        self.add_button = ttk.Button(button_frame, text="Add Files", command=self.add_files, width=15)
        self.add_button.pack(side=tk.LEFT, padx=(0, 10))
        
        # Clear all button
        self.clear_button = ttk.Button(button_frame, text="Clear All", command=self.clear_files, width=15)
        self.clear_button.pack(side=tk.LEFT, padx=(0, 10))
        
        # Merge button
        self.merge_button = ttk.Button(button_frame, text="Merge to PDF", command=self.merge_documents, width=15)
        self.merge_button.pack(side=tk.LEFT)
        
        # Files listbox with scrollbar
        list_frame = ttk.Frame(self.parent_frame)
        list_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)
        
        # Create Treeview for better display
        columns = ('Order', 'Filename', 'Full Path', 'Type', 'Status')
        self.files_tree = ttk.Treeview(list_frame, columns=columns, show='headings', height=10)
        
        # Configure columns
        self.files_tree.heading('Order', text='Order')
        self.files_tree.heading('Filename', text='Filename')
        self.files_tree.heading('Full Path', text='Full Path')
        self.files_tree.heading('Type', text='Type')
        self.files_tree.heading('Status', text='Status')
        
        self.files_tree.column('Order', width=60, anchor='center')
        self.files_tree.column('Filename', width=200, anchor='w')
        self.files_tree.column('Full Path', width=300, anchor='w')
        self.files_tree.column('Type', width=150, anchor='w')
        self.files_tree.column('Status', width=120, anchor='center')
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.files_tree.yview)
        self.files_tree.configure(yscrollcommand=scrollbar.set)
        
        # Bind double-click event to open files
        self.files_tree.bind('<Double-1>', self.on_file_double_click)
        
        # Bind right-click event for context menu
        self.files_tree.bind('<Button-3>', self.on_file_right_click)
        
        # Create context menu
        self.context_menu = tk.Menu(self.parent_frame, tearoff=0)
        self.context_menu.add_command(label="Open File", command=self.open_selected_file)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="Remove from List", command=self.remove_selected)
        
        self.files_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # Remove selected button
        self.remove_button = ttk.Button(self.parent_frame, text="Remove Selected", command=self.remove_selected, width=15)
        self.remove_button.grid(row=3, column=0, columnspan=3, pady=(0, 10))
        
        # Output file name frame
        output_frame = ttk.LabelFrame(self.parent_frame, text="Output File", padding="5")
        output_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        output_frame.columnconfigure(1, weight=1)
        
        # Output file name label and entry
        ttk.Label(output_frame, text="Output File Name:").grid(row=0, column=0, padx=(0, 10), sticky='w')
        self.output_filename = tk.StringVar(value="merged_document.pdf")
        self.output_entry = ttk.Entry(output_frame, textvariable=self.output_filename, width=40)
        self.output_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 10))
        
        # Bind auto-save to output filename changes
        self.output_filename.trace('w', self.on_output_filename_change)
        
        # Browse button for output file
        self.browse_output_button = ttk.Button(output_frame, text="Browse", command=self.browse_output_file, width=10)
        self.browse_output_button.grid(row=0, column=2)
        
        # Page size normalization option
        self.normalize_page_sizes = tk.BooleanVar(value=False)
        normalize_checkbox = ttk.Checkbutton(
            output_frame, 
            text="Normalize page sizes (make all pages the same size)", 
            variable=self.normalize_page_sizes,
            command=self.on_normalize_option_change
        )
        normalize_checkbox.grid(row=1, column=0, columnspan=3, sticky='w', pady=(5, 0))
        
        # Excel sheet ignore option
        excel_ignore_frame = ttk.Frame(output_frame)
        excel_ignore_frame.grid(row=2, column=0, columnspan=3, sticky='w', pady=(5, 0))
        
        self.ignore_excel_sheets = tk.BooleanVar(value=False)
        ignore_checkbox = ttk.Checkbutton(
            excel_ignore_frame,
            text="Ignore Excel sheets containing:",
            variable=self.ignore_excel_sheets,
            command=self.on_excel_ignore_option_change
        )
        ignore_checkbox.pack(side=tk.LEFT, padx=(0, 5))
        
        self.excel_ignore_substring = tk.StringVar(value="")
        ignore_entry = ttk.Entry(excel_ignore_frame, textvariable=self.excel_ignore_substring, width=20)
        ignore_entry.pack(side=tk.LEFT, padx=(0, 5))
        
        # Bind auto-save to ignore substring changes
        self.excel_ignore_substring.trace('w', self.on_excel_ignore_substring_change)
        
        ttk.Label(excel_ignore_frame, text="(case-insensitive)", font=("Arial", 8), foreground="gray").pack(side=tk.LEFT)
        
        # Progress bar
        self.progress = ttk.Progressbar(self.parent_frame, mode='indeterminate')
        self.progress.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Status label
        self.status_label = ttk.Label(self.parent_frame, text="Ready to add files", font=("Arial", 9))
        self.status_label.grid(row=6, column=0, columnspan=3)
        
        # Store reference to main app for accessing supported extensions
        self.main_app = None
        
    def set_main_app(self, main_app):
        """Set reference to main app for accessing shared resources"""
        self.main_app = main_app
        
    def add_files(self):
        if not self.main_app:
            return
            
        filetypes = [
            ("All supported files", "*.xlsx *.xls *.xlsm *.doc *.docx *.pdf"),
            ("Excel files", "*.xlsx *.xls *.xlsm"),
            ("Word files", "*.doc *.docx"),
            ("PDF files", "*.pdf"),
            ("All files", "*.*")
        ]
        
        files = filedialog.askopenfilenames(
            title="Select files to merge",
            filetypes=filetypes
        )
        
        for file_path in files:
            if file_path not in [f[1] for f in self.selected_files]:
                self.add_file_to_list(file_path)
                
    def add_file_to_list(self, file_path):
        if not self.main_app:
            return
            
        file_ext = Path(file_path).suffix.lower()
        if file_ext in self.main_app.supported_extensions:
            order = len(self.selected_files) + 1
            filename = Path(file_path).name
            file_type = self.main_app.supported_extensions[file_ext]
            
            self.selected_files.append((order, file_path, filename, file_type))
            self.files_tree.insert('', 'end', values=(order, filename, file_path, file_type, 'Ready'))

            # If this is the first file added, set default output to same name/path (.pdf)
            if len(self.selected_files) == 1:
                first_path = Path(file_path)
                proposed_output_name = f"{first_path.stem}.pdf"
                proposed_output_path = str(first_path.with_suffix('.pdf'))

                # Only auto-set if user hasn't chosen a custom path yet
                has_custom_output_path = hasattr(self, 'output_file_path') and bool(getattr(self, 'output_file_path', None))
                # Also avoid overriding if user already changed the name from the initial default
                current_name = self.output_filename.get().strip()
                is_default_name = (current_name.lower() == 'merged_document.pdf')

                if not has_custom_output_path and is_default_name:
                    if first_path.suffix.lower() == '.pdf':
                        # Confirm overwrite when first document is already a PDF
                        confirm = messagebox.askyesno(
                            "Confirm Overwrite",
                            "The first document is already a PDF.\n\n"
                            f"Setting the output to '{proposed_output_name}' in the same folder will overwrite the original file.\n\n"
                            "Do you want to overwrite it?"
                        )
                        if not confirm:
                            # Do not change defaults if user declines
                            pass
                        else:
                            self.output_filename.set(proposed_output_name)
                            self.output_file_path = proposed_output_path
                    else:
                        self.output_filename.set(proposed_output_name)
                        self.output_file_path = proposed_output_path
            
            # Trigger auto-save
            if self.main_app:
                self.main_app.auto_save_session()
            
    def remove_selected(self):
        selected_item = self.files_tree.selection()
        if selected_item:
            item_values = self.files_tree.item(selected_item[0])['values']
            file_path = item_values[2]  # Full path is now in column 2
            
            # Remove from selected_files
            self.selected_files = [(order, path, name, type_) for order, path, name, type_ in self.selected_files if path != file_path]
            
            # Remove from treeview
            self.files_tree.delete(selected_item[0])
            
            # Reorder remaining items
            self.reorder_files()
            
            # Trigger auto-save
            if self.main_app:
                self.main_app.auto_save_session()
            
    def reorder_files(self):
        # Clear treeview
        for item in self.files_tree.get_children():
            self.files_tree.delete(item)
            
        # Re-add with new order
        for i, (_, file_path, filename, file_type) in enumerate(self.selected_files, 1):
            self.selected_files[i-1] = (i, file_path, filename, file_type)
            self.files_tree.insert('', 'end', values=(i, filename, file_path, file_type, 'Ready'))
            
    def clear_files(self):
        self.selected_files.clear()
        for item in self.files_tree.get_children():
            self.files_tree.delete(item)
            
        # Trigger auto-save
        if self.main_app:
            self.main_app.auto_save_session()
            
    def browse_output_file(self):
        """Browse for output file location"""
        output_file = filedialog.asksaveasfilename(
            title="Save merged PDF as",
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
            initialfile=self.output_filename.get()
        )
        
        if output_file:
            self.output_filename.set(Path(output_file).name)
            self.output_file_path = output_file
            
    def on_output_filename_change(self, *args):
        """Callback for output filename changes"""
        # Trigger auto-save
        if self.main_app:
            self.main_app.auto_save_session()
            
    def on_normalize_option_change(self):
        """Callback for normalize page sizes option changes"""
        # Trigger auto-save
        if self.main_app:
            self.main_app.auto_save_session()
            
    def on_excel_ignore_option_change(self):
        """Callback for Excel ignore sheets option changes"""
        # Trigger auto-save
        if self.main_app:
            self.main_app.auto_save_session()
            
    def on_excel_ignore_substring_change(self, *args):
        """Callback for Excel ignore substring changes"""
        # Trigger auto-save
        if self.main_app:
            self.main_app.auto_save_session()
            
    def normalize_pdf_page_sizes(self, pdf_files):
        """Normalize page sizes in PDF files to match the largest page dimensions"""
        if not pdf_files:
            return pdf_files
            
        # Find maximum dimensions across all pages
        max_width = 0
        max_height = 0
        
        for pdf_file in pdf_files:
            try:
                reader = PdfReader(pdf_file)
                for page in reader.pages:
                    width = float(page.mediabox.width)
                    height = float(page.mediabox.height)
                    max_width = max(max_width, width)
                    max_height = max(max_height, height)
            except Exception as e:
                print(f"Error reading PDF {pdf_file}: {e}")
                continue
                
        if max_width == 0 or max_height == 0:
            return pdf_files  # Return original files if we can't determine dimensions
            
        # Create normalized PDFs
        normalized_files = []
        for pdf_file in pdf_files:
            try:
                reader = PdfReader(pdf_file)
                writer = PdfWriter()
                
                for page in reader.pages:
                    width = float(page.mediabox.width)
                    height = float(page.mediabox.height)
                    
                    # Calculate scale factor to fit within max dimensions while preserving aspect ratio
                    scale_x = max_width / width
                    scale_y = max_height / height
                    scale = min(scale_x, scale_y)  # Use smaller scale to fit within bounds
                    
                    # Scale the page
                    page.scale_by(scale)
                    writer.add_page(page)
                
                # Create temporary file for normalized PDF
                temp_normalized = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
                temp_normalized.close()
                
                with open(temp_normalized.name, 'wb') as output_file:
                    writer.write(output_file)
                    
                normalized_files.append(temp_normalized.name)
                
            except Exception as e:
                print(f"Error normalizing PDF {pdf_file}: {e}")
                normalized_files.append(pdf_file)  # Use original if normalization fails
                
        return normalized_files
            
    def merge_documents(self):
        if not self.selected_files:
            messagebox.showwarning("No Files", "Please add files to merge first.")
            return
            
        # Get output file path from entry field
        output_filename = self.output_filename.get().strip()
        if not output_filename:
            messagebox.showwarning("No Output Name", "Please specify an output filename.")
            return
            
        # Ensure it has .pdf extension
        if not output_filename.lower().endswith('.pdf'):
            output_filename += '.pdf'
            self.output_filename.set(output_filename)
            
        # Use current directory or get from saved path
        if hasattr(self, 'output_file_path') and self.output_file_path:
            output_dir = os.path.dirname(self.output_file_path)
        else:
            output_dir = os.getcwd()
            
        output_file = os.path.join(output_dir, output_filename)
            
        # Run merge in separate thread to avoid freezing GUI
        thread = threading.Thread(target=self.perform_merge, args=(output_file,))
        thread.daemon = True
        thread.start()
            
    def perform_merge(self, output_file):
        self.main_app.root.after(0, lambda: self.progress.start())
        self.main_app.root.after(0, lambda: self.status_label.config(text="Converting documents..."))
        
        try:
            merger = PdfMerger()
            temp_files = []
            pdf_files = []
            failed_file = None
            
            # Get ignore substring if option is enabled
            ignore_substring = None
            if self.ignore_excel_sheets.get():
                ignore_substring = self.excel_ignore_substring.get().strip()
                if not ignore_substring:
                    ignore_substring = None
            
            # Convert all files to PDF first
            for order, file_path, filename, file_type in self.selected_files:
                self.main_app.root.after(0, lambda f=filename: self.update_file_status(f, 'Converting...'))
                
                pdf_path = self.main_app.convert_to_pdf(file_path, file_type, ignore_substring)
                if pdf_path:
                    pdf_files.append(pdf_path)
                    temp_files.append(pdf_path)
                    self.main_app.root.after(0, lambda f=filename: self.update_file_status(f, 'Converted'))
                else:
                    # Conversion failed - cancel the entire batch
                    failed_file = filename
                    self.main_app.root.after(0, lambda f=filename: self.update_file_status(f, 'Failed'))
                    break
            
            # If any conversion failed, cancel the operation
            if failed_file:
                # Clean up temp files created so far
                for temp_file in temp_files:
                    try:
                        os.remove(temp_file)
                    except:
                        pass
                
                self.main_app.root.after(0, lambda: self.progress.stop())
                self.main_app.root.after(0, lambda: self.status_label.config(text="Merge cancelled - conversion failed"))
                self.main_app.root.after(0, lambda f=failed_file: messagebox.showerror("Conversion Failed", 
                    f"Failed to convert '{f}' to PDF.\n\nBatch operation cancelled. Please fix the file and try again."))
                merger.close()
                return
            
            # Normalize page sizes if option is enabled
            if self.normalize_page_sizes.get() and pdf_files:
                self.main_app.root.after(0, lambda: self.status_label.config(text="Normalizing page sizes..."))
                normalized_files = self.normalize_pdf_page_sizes(pdf_files)
                temp_files.extend(normalized_files)
                
                # Use normalized files for merging
                for pdf_file in normalized_files:
                    merger.append(pdf_file)
            else:
                # Use original PDF files
                for pdf_file in pdf_files:
                    merger.append(pdf_file)
                    
            # Write merged PDF
            self.main_app.root.after(0, lambda: self.status_label.config(text="Creating final PDF..."))
            merger.write(output_file)
            merger.close()
            
            # Clean up temp files
            for temp_file in temp_files:
                try:
                    os.remove(temp_file)
                except:
                    pass
                    
            self.main_app.root.after(0, lambda: self.progress.stop())
            self.main_app.root.after(0, lambda: self.status_label.config(text="Merge completed successfully!"))
            self.main_app.root.after(0, lambda: self.main_app.ask_to_open_file(output_file))
            
        except Exception as e:
            self.main_app.root.after(0, lambda: self.progress.stop())
            self.main_app.root.after(0, lambda: self.status_label.config(text="Error occurred during merge"))
            self.main_app.root.after(0, lambda: messagebox.showerror("Error", f"An error occurred: {str(e)}"))
            
    def update_file_status(self, filename, status):
        for item in self.files_tree.get_children():
            values = self.files_tree.item(item)['values']
            # Check if the filename matches (now in column 1)
            if values[1] == filename:
                self.files_tree.set(item, 'Status', status)
                break
                
    def on_file_double_click(self, event):
        """Handle double-click on file in the list to open it"""
        # Get the selected item
        selected_item = self.files_tree.selection()
        if not selected_item:
            return
            
        # Get the file path from the selected item
        item_values = self.files_tree.item(selected_item[0])['values']
        file_path = item_values[2]  # Full path is in column 2
        
        # Check if file exists
        if not os.path.exists(file_path):
            messagebox.showerror("File Not Found", f"The file does not exist:\n{file_path}")
            return
            
        try:
            # Open the file with the default system application
            os.startfile(file_path)
        except Exception as e:
            messagebox.showerror("Error Opening File", f"Could not open the file:\n{str(e)}")
            
    def on_file_right_click(self, event):
        """Handle right-click on file in the list to show context menu"""
        # Select the item under the cursor
        item = self.files_tree.identify_row(event.y)
        if item:
            self.files_tree.selection_set(item)
            # Show context menu at cursor position
            self.context_menu.tk_popup(event.x_root, event.y_root)
            
    def open_selected_file(self):
        """Open the currently selected file"""
        selected_item = self.files_tree.selection()
        if not selected_item:
            return
            
        # Get the file path from the selected item
        item_values = self.files_tree.item(selected_item[0])['values']
        file_path = item_values[2]  # Full path is in column 2
        
        # Check if file exists
        if not os.path.exists(file_path):
            messagebox.showerror("File Not Found", f"The file does not exist:\n{file_path}")
            return
            
        try:
            # Open the file with the default system application
            os.startfile(file_path)
        except Exception as e:
            messagebox.showerror("Error Opening File", f"Could not open the file:\n{str(e)}")
                
    def get_save_data(self):
        """Get data to save for this tab"""
        return {
            'output_filename': self.output_filename.get(),
            'output_file_path': getattr(self, 'output_file_path', None),
            'selected_files': self.selected_files,
            'normalize_page_sizes': self.normalize_page_sizes.get(),
            'ignore_excel_sheets': self.ignore_excel_sheets.get(),
            'excel_ignore_substring': self.excel_ignore_substring.get()
        }
        
    def load_save_data(self, data):
        """Load data from saved configuration"""
        # Clear current files
        self.selected_files.clear()
        for item in self.files_tree.get_children():
            self.files_tree.delete(item)
            
        # Load output filename
        if 'output_filename' in data:
            self.output_filename.set(data['output_filename'])
            
        # Load output file path
        if 'output_file_path' in data and data['output_file_path']:
            self.output_file_path = data['output_file_path']
            
        # Load normalize page sizes option
        if 'normalize_page_sizes' in data:
            self.normalize_page_sizes.set(data['normalize_page_sizes'])
            
        # Load Excel ignore sheets option
        if 'ignore_excel_sheets' in data:
            self.ignore_excel_sheets.set(data['ignore_excel_sheets'])
            
        # Load Excel ignore substring
        if 'excel_ignore_substring' in data:
            self.excel_ignore_substring.set(data['excel_ignore_substring'])
            
        # Load selected files
        if 'selected_files' in data:
            for order, file_path, filename, file_type in data['selected_files']:
                # Check if file still exists
                if os.path.exists(file_path):
                    self.selected_files.append((order, file_path, filename, file_type))
                    self.files_tree.insert('', 'end', values=(order, filename, file_path, file_type, 'Ready'))
                else:
                    # File doesn't exist, add with 'File Missing' status
                    self.selected_files.append((order, file_path, filename, file_type))
                    self.files_tree.insert('', 'end', values=(order, filename, file_path, file_type, 'File Missing'))

class DocumentMerger:
    def __init__(self, root):
        self.root = root
        self.root.title("Document Merger - Multi-Tab")
        self.root.geometry("800x700")
        self.root.minsize(750, 650)
        
        # Supported file extensions
        self.supported_extensions = {
            '.xlsx': 'Excel Workbook',
            '.xls': 'Excel Workbook (Legacy)',
            '.xlsm': 'Excel Macro-Enabled Workbook',
            '.doc': 'Word Document (Legacy)',
            '.docx': 'Word Document',
            '.pdf': 'PDF Document'
        }
        
        # Dictionary to store tab contents
        self.tabs = {}
        self.tab_counter = 1
        
        # Auto-save file path
        self.auto_save_file = os.path.join(os.getcwd(), ".document_merger_autosave.json")
        
        self.setup_ui()
        
        # Load last session on startup
        self.load_last_session()
        
    def setup_ui(self):
        # Create menu bar
        self.create_menu_bar()
        
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(1, weight=1)  # Notebook row gets the weight
        
        # Top frame for Add New Tab button (top right)
        top_frame = ttk.Frame(main_frame)
        top_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        top_frame.columnconfigure(0, weight=1)  # Push button to the right
        
        add_tab_button = ttk.Button(top_frame, text="Add New Tab", command=self.add_new_tab, width=15)
        add_tab_button.grid(row=0, column=1, sticky=tk.E)  # Position on the right
        
        # Create notebook for tabs
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Add first tab
        self.add_new_tab()
        
        # Tab management buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=2, column=0, pady=(10, 0), sticky=(tk.W, tk.E))
        
        remove_tab_button = ttk.Button(button_frame, text="Remove Current Tab", command=self.remove_current_tab, width=15)
        remove_tab_button.pack(side=tk.LEFT, padx=(0, 10))
        
        # Merge all button
        merge_all_button = ttk.Button(button_frame, text="Merge All Tabs", command=self.merge_all_tabs, width=15)
        merge_all_button.pack(side=tk.LEFT)
        
    def create_menu_bar(self):
        """Create the menu bar with File menu"""
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        # File menu
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        
        # File menu items
        file_menu.add_command(label="Save All Tabs", command=self.save_all_tabs, accelerator="Ctrl+S")
        file_menu.add_command(label="Load All Tabs", command=self.load_all_tabs, accelerator="Ctrl+O")
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.root.quit, accelerator="Ctrl+Q")
        
        # Excel Macros menu
        macros_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Excel Macros", menu=macros_menu)
        
        # Add macro files to menu
        self.populate_macros_menu(macros_menu)
        
        # Setup keyboard shortcuts
        self.root.bind('<Control-s>', lambda e: self.save_all_tabs())
        self.root.bind('<Control-o>', lambda e: self.load_all_tabs())
        self.root.bind('<Control-q>', lambda e: self.root.quit())
        
        # Focus on the main window for keyboard shortcuts
        self.root.focus_set()
        
    def populate_macros_menu(self, macros_menu):
        """Populate the Excel Macros menu with available macro files"""
        macros_folder = os.path.join(os.getcwd(), "Excel Macros")
        
        if not os.path.exists(macros_folder):
            macros_menu.add_command(label="No macros folder found", state="disabled")
            return
            
        # Get all .vba, .bas, .txt files from the Excel Macros folder
        macro_files = []
        for file in os.listdir(macros_folder):
            if file.lower().endswith(('.vba', '.bas', '.txt')):
                macro_files.append(file)
                
        if not macro_files:
            macros_menu.add_command(label="No macro files found", state="disabled")
            return
            
        # Sort files alphabetically
        macro_files.sort()
        
        # Add each macro file as a menu item
        for macro_file in macro_files:
            macros_menu.add_command(
                label=macro_file,
                command=lambda f=macro_file: self.copy_macro_to_clipboard(f)
            )
            
    def copy_macro_to_clipboard(self, macro_filename):
        """Copy the content of a macro file to the clipboard"""
        try:
            macros_folder = os.path.join(os.getcwd(), "Excel Macros")
            macro_path = os.path.join(macros_folder, macro_filename)
            
            if not os.path.exists(macro_path):
                messagebox.showerror("Error", f"Macro file not found: {macro_filename}")
                return
                
            # Read the macro file content
            with open(macro_path, 'r', encoding='utf-8') as f:
                macro_content = f.read()
                
            # Copy to clipboard
            pyperclip.copy(macro_content)
            
            # Show success message
            messagebox.showinfo("Success", f"Macro '{macro_filename}' copied to clipboard!")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to copy macro to clipboard:\n{str(e)}")
        
    def auto_save_session(self):
        """Automatically save current session to auto-save file"""
        try:
            if not self.tabs:
                print("DEBUG: No tabs to auto-save")
                return
                
            # Check if any tab has files selected
            has_files = False
            for tab_content in self.tabs.values():
                if tab_content.selected_files:
                    has_files = True
                    break
                    
            if not has_files:
                print("DEBUG: No files selected in any tab, skipping auto-save")
                return
                
            print(f"DEBUG: Auto-saving session to: {self.auto_save_file}")
            # Collect data from all tabs
            save_data = {
                'version': '1.0',
                'saved_date': datetime.now().isoformat(),
                'auto_save': True,
                'tabs': {}
            }
            
            for tab_name, tab_content in self.tabs.items():
                save_data['tabs'][tab_name] = tab_content.get_save_data()
                
            print(f"DEBUG: Auto-save data: {save_data}")
            # Save to auto-save file
            with open(self.auto_save_file, 'w', encoding='utf-8') as f:
                json.dump(save_data, f, indent=2, ensure_ascii=False)
                
            print("DEBUG: Auto-save completed successfully")
                
        except Exception as e:
            # Silently fail for auto-save to avoid interrupting user workflow
            print(f"DEBUG: Auto-save failed with exception: {e}")
            pass
            
    def load_last_session(self):
        """Load the last session from auto-save file"""
        try:
            print(f"DEBUG: Looking for auto-save file at: {self.auto_save_file}")
            if not os.path.exists(self.auto_save_file):
                print("DEBUG: Auto-save file not found")
                return
                
            print("DEBUG: Auto-save file found, loading...")
            with open(self.auto_save_file, 'r', encoding='utf-8') as f:
                save_data = json.load(f)
                
            print(f"DEBUG: Loaded save data: {save_data}")
            # Validate file format
            if 'tabs' not in save_data or not save_data.get('auto_save', False):
                print("DEBUG: Invalid auto-save file format or not an auto-save file")
                return
                
            print("DEBUG: Clearing existing tabs...")
            # Clear existing tabs first
            for tab_name in list(self.tabs.keys()):
                current_tab = self.notebook.select()
                if current_tab:
                    self.notebook.forget(current_tab)
                    del self.tabs[tab_name]
                    
            # Reset tab counter
            self.tab_counter = 1
            
            print(f"DEBUG: Loading {len(save_data['tabs'])} tabs from auto-save data...")
            # Load tabs from auto-save data
            for tab_name, tab_data in save_data['tabs'].items():
                print(f"DEBUG: Loading tab: {tab_name}")
                # Create new tab
                tab_frame = ttk.Frame(self.notebook)
                
                # Create tab content
                tab_content = TabContent(tab_frame)
                tab_content.set_main_app(self)
                
                # Load the saved data
                tab_content.load_save_data(tab_data)
                
                # Add tab to notebook
                self.notebook.add(tab_frame, text=tab_name)
                self.tabs[tab_name] = tab_content
                
                # Update tab counter
                self.tab_counter = max(self.tab_counter, int(tab_name.split()[-1]) + 1)
                
            # Select first tab
            if self.tabs:
                first_tab = list(self.tabs.keys())[0]
                for tab_id in self.notebook.tabs():
                    if self.notebook.tab(tab_id, "text") == first_tab:
                        self.notebook.select(tab_id)
                        break
                        
            print("DEBUG: Auto-load completed successfully")
                        
        except Exception as e:
            # Silently fail for auto-load to avoid interrupting startup
            print(f"DEBUG: Auto-load failed with exception: {e}")
            pass
            
    def cleanup_auto_save(self):
        """Clean up auto-save file on exit"""
        # Don't delete the auto-save file so it can be loaded on next startup
        pass
        
    def add_new_tab(self):
        """Add a new tab to the notebook"""
        tab_frame = ttk.Frame(self.notebook)
        tab_name = f"Batch {self.tab_counter}"
        
        # Create tab content
        tab_content = TabContent(tab_frame)
        tab_content.set_main_app(self)
        
        # Add tab to notebook
        self.notebook.add(tab_frame, text=tab_name)
        self.tabs[tab_name] = tab_content
        
        # Switch to new tab
        self.notebook.select(tab_frame)
        
        self.tab_counter += 1
        
        # Trigger auto-save
        self.auto_save_session()
        
    def remove_current_tab(self):
        """Remove the currently selected tab"""
        current_tab = self.notebook.select()
        if current_tab:
            tab_name = self.notebook.tab(current_tab, "text")
            
            # Don't remove if it's the last tab
            if len(self.tabs) > 1:
                self.notebook.forget(current_tab)
                del self.tabs[tab_name]
                
                # Trigger auto-save
                self.auto_save_session()
            else:
                messagebox.showinfo("Info", "Cannot remove the last tab. At least one tab must remain.")
                
    def save_all_tabs(self):
        """Save all tabs configuration to a JSON file"""
        if not self.tabs:
            messagebox.showwarning("No Tabs", "No tabs to save.")
            return
            
        # Collect data from all tabs
        save_data = {
            'version': '1.0',
            'saved_date': datetime.now().isoformat(),
            'tabs': {}
        }
        
        for tab_name, tab_content in self.tabs.items():
            save_data['tabs'][tab_name] = tab_content.get_save_data()
            
        # Ask user where to save
        file_path = filedialog.asksaveasfilename(
            title="Save Tab Configuration",
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
            initialfile=f"document_merger_config_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        )
        
        if file_path:
            try:
                with open(file_path, 'w', encoding='utf-8') as f:
                    json.dump(save_data, f, indent=2, ensure_ascii=False)
                messagebox.showinfo("Success", f"Configuration saved successfully to:\n{file_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save configuration:\n{str(e)}")
                
    def load_all_tabs(self):
        """Load all tabs configuration from a JSON file"""
        file_path = filedialog.askopenfilename(
            title="Load Tab Configuration",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        
        if not file_path:
            return
            
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                save_data = json.load(f)
                
            # Validate file format
            if 'tabs' not in save_data:
                messagebox.showerror("Error", "Invalid configuration file format.")
                return
                
            # Clear existing tabs
            for tab_name in list(self.tabs.keys()):
                current_tab = self.notebook.select()
                if current_tab:
                    self.notebook.forget(current_tab)
                    del self.tabs[tab_name]
                    
            # Reset tab counter
            self.tab_counter = 1
            
            # Load tabs from saved data
            for tab_name, tab_data in save_data['tabs'].items():
                # Create new tab
                tab_frame = ttk.Frame(self.notebook)
                
                # Create tab content
                tab_content = TabContent(tab_frame)
                tab_content.set_main_app(self)
                
                # Load the saved data
                tab_content.load_save_data(tab_data)
                
                # Add tab to notebook
                self.notebook.add(tab_frame, text=tab_name)
                self.tabs[tab_name] = tab_content
                
                # Update tab counter
                self.tab_counter = max(self.tab_counter, int(tab_name.split()[-1]) + 1)
                
            # Select first tab
            if self.tabs:
                first_tab = list(self.tabs.keys())[0]
                for tab_id in self.notebook.tabs():
                    if self.notebook.tab(tab_id, "text") == first_tab:
                        self.notebook.select(tab_id)
                        break
                        
            messagebox.showinfo("Success", f"Configuration loaded successfully from:\n{file_path}")
            
        except FileNotFoundError:
            messagebox.showerror("Error", "Configuration file not found.")
        except json.JSONDecodeError:
            messagebox.showerror("Error", "Invalid JSON format in configuration file.")
        except Exception as e:
                         messagebox.showerror("Error", f"Failed to load configuration:\n{str(e)}")
                
    def merge_all_tabs(self):
        """Merge all tabs at once"""
        if not self.tabs:
            messagebox.showwarning("No Tabs", "No tabs to merge.")
            return
            
        # Check if any tab has files
        tabs_with_files = []
        for tab_name, tab_content in self.tabs.items():
            if tab_content.selected_files:
                tabs_with_files.append((tab_name, tab_content))
                
        if not tabs_with_files:
            messagebox.showwarning("No Files", "No files to merge in any tab.")
            return
            
        # Validate output filenames
        invalid_tabs = []
        for tab_name, tab_content in tabs_with_files:
            output_filename = tab_content.output_filename.get().strip()
            if not output_filename:
                invalid_tabs.append(tab_name)
                
        if invalid_tabs:
            messagebox.showwarning("Missing Output Names", 
                                 f"The following tabs are missing output filenames:\n{', '.join(invalid_tabs)}")
            return
            
        # Confirm merge all
        result = messagebox.askyesno("Merge All Tabs", 
                                   f"Merge all {len(tabs_with_files)} tabs?\n\nThis will create {len(tabs_with_files)} PDF files.")
        if not result:
            return
            
        # Run merge all in separate thread
        thread = threading.Thread(target=self.perform_merge_all, args=(tabs_with_files,))
        thread.daemon = True
        thread.start()
        
    def perform_merge_all(self, tabs_with_files):
        """Perform merge for all tabs"""
        completed_files = []
        failed_tabs = []
        
        total_tabs = len(tabs_with_files)
        
        for i, (tab_name, tab_content) in enumerate(tabs_with_files, 1):
            try:
                # Update status
                self.root.after(0, lambda: self.update_merge_all_status(f"Processing {tab_name} ({i}/{total_tabs})..."))
                
                # Get output filename
                output_filename = tab_content.output_filename.get().strip()
                if not output_filename.lower().endswith('.pdf'):
                    output_filename += '.pdf'
                    
                # Use current directory or get from saved path
                if hasattr(tab_content, 'output_file_path') and tab_content.output_file_path:
                    output_dir = os.path.dirname(tab_content.output_file_path)
                else:
                    output_dir = os.getcwd()
                    
                output_file = os.path.join(output_dir, output_filename)
                
                # Perform merge for this tab
                merger = PdfMerger()
                temp_files = []
                pdf_files = []
                failed_file = None
                
                # Get ignore substring if option is enabled for this tab
                ignore_substring = None
                if tab_content.ignore_excel_sheets.get():
                    ignore_substring = tab_content.excel_ignore_substring.get().strip()
                    if not ignore_substring:
                        ignore_substring = None
                
                # Convert all files to PDF first
                for order, file_path, filename, file_type in tab_content.selected_files:
                    # Update file status
                    self.root.after(0, lambda f=filename, t=tab_name: self.update_tab_file_status(f, t, 'Converting...'))
                    
                    pdf_path = self.convert_to_pdf(file_path, file_type, ignore_substring)
                    if pdf_path:
                        pdf_files.append(pdf_path)
                        temp_files.append(pdf_path)
                        self.root.after(0, lambda f=filename, t=tab_name: self.update_tab_file_status(f, t, 'Converted'))
                    else:
                        # Conversion failed - cancel this batch
                        failed_file = filename
                        self.root.after(0, lambda f=filename, t=tab_name: self.update_tab_file_status(f, t, 'Failed'))
                        break
                
                # If any conversion failed, cancel this batch and continue with next tab
                if failed_file:
                    # Clean up temp files created so far
                    for temp_file in temp_files:
                        try:
                            os.remove(temp_file)
                        except:
                            pass
                    
                    merger.close()
                    failed_tabs.append((tab_name, f"Conversion failed for '{failed_file}'"))
                    self.root.after(0, lambda t=tab_name, f=failed_file: self.update_merge_all_status(f"Cancelled {t}: failed to convert '{f}'"))
                    continue
                
                # Normalize page sizes if option is enabled
                if tab_content.normalize_page_sizes.get() and pdf_files:
                    self.root.after(0, lambda: self.update_merge_all_status(f"Normalizing page sizes for {tab_name}..."))
                    normalized_files = tab_content.normalize_pdf_page_sizes(pdf_files)
                    temp_files.extend(normalized_files)
                    
                    # Use normalized files for merging
                    for pdf_file in normalized_files:
                        merger.append(pdf_file)
                else:
                    # Use original PDF files
                    for pdf_file in pdf_files:
                        merger.append(pdf_file)
                        
                # Write merged PDF
                merger.write(output_file)
                merger.close()
                
                # Clean up temp files
                for temp_file in temp_files:
                    try:
                        os.remove(temp_file)
                    except:
                        pass
                        
                completed_files.append(output_file)
                self.root.after(0, lambda f=output_file: self.update_merge_all_status(f"Completed: {os.path.basename(f)}"))
                
            except Exception as e:
                failed_tabs.append((tab_name, str(e)))
                self.root.after(0, lambda t=tab_name, e=str(e): self.update_merge_all_status(f"Failed {t}: {e}"))
                
        # Show completion message
        self.root.after(0, lambda: self.show_merge_all_completion(completed_files, failed_tabs))
        
    def update_merge_all_status(self, message):
        """Update status for merge all operation"""
        # Update the status of the first tab (or create a global status)
        if self.tabs:
            first_tab = list(self.tabs.values())[0]
            first_tab.status_label.config(text=message)
            
    def update_tab_file_status(self, filename, tab_name, status):
        """Update file status in specific tab"""
        if tab_name in self.tabs:
            tab_content = self.tabs[tab_name]
            tab_content.update_file_status(filename, status)
            
    def show_merge_all_completion(self, completed_files, failed_tabs):
        """Show completion message for merge all operation"""
        message = f"Merge All Complete!\n\n"
        
        if completed_files:
            message += f"Successfully created {len(completed_files)} files:\n"
            for file_path in completed_files:
                message += f"• {os.path.basename(file_path)}\n"
            message += "\n"
            
        if failed_tabs:
            message += f"Failed to merge {len(failed_tabs)} tabs:\n"
            for tab_name, error in failed_tabs:
                message += f"• {tab_name}: {error}\n"
                
        if completed_files:
            result = messagebox.askyesno("Merge All Complete", 
                                       f"{message}\nWould you like to open all completed PDF files?")
            if result:
                for file_path in completed_files:
                    try:
                        os.startfile(file_path)
                    except Exception as e:
                        messagebox.showerror("Error Opening File", 
                                           f"Could not open {os.path.basename(file_path)}:\n{str(e)}")
        else:
            messagebox.showinfo("Merge All Complete", message)
                
    def convert_to_pdf(self, file_path, file_type, ignore_substring=None):
        file_ext = Path(file_path).suffix.lower()
        
        if file_ext == '.pdf':
            return file_path
        elif file_ext in ['.xlsx', '.xls', '.xlsm']:
            return self.convert_excel_to_pdf(file_path, ignore_substring)
        elif file_ext in ['.doc', '.docx']:
            return self.convert_word_to_pdf(file_path)
        else:
            return None
            
    def convert_excel_to_pdf(self, file_path, ignore_substring=None):
        excel = None
        created_excel = False
        try:
            # Check if file exists
            if not os.path.exists(file_path):
                print(f"Excel file not found: {file_path}")
                return None
                
            # Create temporary PDF file
            temp_pdf = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
            temp_pdf.close()
            
            # Normalize the file path to handle special characters and encoding
            normalized_path = os.path.abspath(file_path)
            
            # Try to get existing Excel instance first
            try:
                excel = win32com.client.GetActiveObject("Excel.Application")
                # Excel was already running, don't quit it
                created_excel = False
            except pythoncom.com_error:
                # Excel is not running, create a new instance
                excel = win32com.client.Dispatch("Excel.Application")
                created_excel = True
            
            excel.Visible = False
            excel.DisplayAlerts = False
            
            workbook = excel.Workbooks.Open(normalized_path, ReadOnly=True)
            
            # Hide sheets that contain the ignore substring (case-insensitive)
            if ignore_substring and ignore_substring.strip():
                ignore_substring_lower = ignore_substring.strip().lower()
                sheets_to_hide = []
                for sheet in workbook.Sheets:
                    if ignore_substring_lower in sheet.Name.lower():
                        sheets_to_hide.append(sheet)
                
                # Hide the sheets
                for sheet in sheets_to_hide:
                    sheet.Visible = False
            
            workbook.ExportAsFixedFormat(0, temp_pdf.name)  # 0 = PDF format
            
            # Restore sheet visibility (in case Excel instance is reused)
            if ignore_substring and ignore_substring.strip():
                for sheet in workbook.Sheets:
                    if not sheet.Visible:
                        sheet.Visible = True
            
            workbook.Close(SaveChanges=False)
            
            # Only quit Excel if we created it
            if created_excel:
                excel.Quit()
            
            return temp_pdf.name
            
        except Exception as e:
            print(f"Error converting Excel file '{file_path}': {e}")
            # Make sure to quit Excel if we created it and there was an error
            if excel and created_excel:
                try:
                    excel.Quit()
                except:
                    pass
            return None
            
    def convert_word_to_pdf(self, file_path):
        word = None
        created_word = False
        try:
            # Check if file exists
            if not os.path.exists(file_path):
                print(f"Word file not found: {file_path}")
                return None
                
            # Create temporary PDF file
            temp_pdf = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
            temp_pdf.close()
            
            # Normalize the file path to handle special characters and encoding
            normalized_path = os.path.abspath(file_path)
            
            # Try to get existing Word instance first
            try:
                word = win32com.client.GetActiveObject("Word.Application")
                # Word was already running, don't quit it
                created_word = False
            except pythoncom.com_error:
                # Word is not running, create a new instance
                word = win32com.client.Dispatch("Word.Application")
                created_word = True
            
            word.Visible = False
            word.DisplayAlerts = False
            
            doc = word.Documents.Open(normalized_path, ReadOnly=True)
            doc.SaveAs(temp_pdf.name, FileFormat=17)  # 17 = PDF format
            doc.Close(SaveChanges=False)
            
            # Only quit Word if we created it
            if created_word:
                word.Quit()
            
            return temp_pdf.name
            
        except Exception as e:
            print(f"Error converting Word file '{file_path}': {e}")
            # Make sure to quit Word if we created it and there was an error
            if word and created_word:
                try:
                    word.Quit()
                except:
                    pass
            return None
            
    def ask_to_open_file(self, output_file):
        """Ask user if they want to open the merged PDF file"""
        result = messagebox.askyesno(
            "Merge Complete", 
            f"Documents merged successfully!\n\nSaved as: {output_file}\n\nWould you like to open the PDF file now?"
        )
        
        if result:
            try:
                # Use the default system application to open the PDF
                os.startfile(output_file)
            except Exception as e:
                messagebox.showerror(
                    "Error Opening File", 
                    f"Could not open the PDF file:\n{str(e)}\n\nYou can manually open it from:\n{output_file}"
                )

def main():
    root = tk.Tk()
    app = DocumentMerger(root)
    
    # Handle window close event
    def on_closing():
        app.cleanup_auto_save()
        root.destroy()
    
    root.protocol("WM_DELETE_WINDOW", on_closing)
    root.mainloop()

if __name__ == "__main__":
    main()
