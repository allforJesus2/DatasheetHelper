import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import sys
import win32com.client
import pandas as pd
from PyPDF2 import PdfMerger
import tempfile
import threading
from pathlib import Path
import json
from datetime import datetime

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
        columns = ('Order', 'Filename', 'Type', 'Status')
        self.files_tree = ttk.Treeview(list_frame, columns=columns, show='headings', height=10)
        
        # Configure columns
        self.files_tree.heading('Order', text='Order')
        self.files_tree.heading('Filename', text='Filename')
        self.files_tree.heading('Type', text='Type')
        self.files_tree.heading('Status', text='Status')
        
        self.files_tree.column('Order', width=60, anchor='center')
        self.files_tree.column('Filename', width=350, anchor='w')
        self.files_tree.column('Type', width=180, anchor='w')
        self.files_tree.column('Status', width=120, anchor='center')
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.files_tree.yview)
        self.files_tree.configure(yscrollcommand=scrollbar.set)
        
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
        
        # Browse button for output file
        self.browse_output_button = ttk.Button(output_frame, text="Browse", command=self.browse_output_file, width=10)
        self.browse_output_button.grid(row=0, column=2)
        
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
            self.files_tree.insert('', 'end', values=(order, filename, file_type, 'Ready'))
            
    def remove_selected(self):
        selected_item = self.files_tree.selection()
        if selected_item:
            item_values = self.files_tree.item(selected_item[0])['values']
            filename = item_values[1]
            
            # Remove from selected_files
            self.selected_files = [(order, path, name, type_) for order, path, name, type_ in self.selected_files if name != filename]
            
            # Remove from treeview
            self.files_tree.delete(selected_item[0])
            
            # Reorder remaining items
            self.reorder_files()
            
    def reorder_files(self):
        # Clear treeview
        for item in self.files_tree.get_children():
            self.files_tree.delete(item)
            
        # Re-add with new order
        for i, (_, file_path, filename, file_type) in enumerate(self.selected_files, 1):
            self.selected_files[i-1] = (i, file_path, filename, file_type)
            self.files_tree.insert('', 'end', values=(i, filename, file_type, 'Ready'))
            
    def clear_files(self):
        self.selected_files.clear()
        for item in self.files_tree.get_children():
            self.files_tree.delete(item)
            
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
            
            for order, file_path, filename, file_type in self.selected_files:
                self.main_app.root.after(0, lambda f=filename: self.update_file_status(f, 'Converting...'))
                
                pdf_path = self.main_app.convert_to_pdf(file_path, file_type)
                if pdf_path:
                    merger.append(pdf_path)
                    temp_files.append(pdf_path)
                    self.main_app.root.after(0, lambda f=filename: self.update_file_status(f, 'Converted'))
                else:
                    self.main_app.root.after(0, lambda f=filename: self.update_file_status(f, 'Failed'))
                    
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
            if values[1] == filename:
                self.files_tree.set(item, 'Status', status)
                break
                
    def get_save_data(self):
        """Get data to save for this tab"""
        return {
            'output_filename': self.output_filename.get(),
            'output_file_path': getattr(self, 'output_file_path', None),
            'selected_files': self.selected_files
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
            
        # Load selected files
        if 'selected_files' in data:
            for order, file_path, filename, file_type in data['selected_files']:
                # Check if file still exists
                if os.path.exists(file_path):
                    self.selected_files.append((order, file_path, filename, file_type))
                    self.files_tree.insert('', 'end', values=(order, filename, file_type, 'Ready'))
                else:
                    # File doesn't exist, add with 'File Missing' status
                    self.selected_files.append((order, file_path, filename, file_type))
                    self.files_tree.insert('', 'end', values=(order, filename, file_type, 'File Missing'))

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
        
        self.setup_ui()
        
    def setup_ui(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(1, weight=1)
        
        # Title
        title_label = ttk.Label(main_frame, text="Document Merger - Multi-Tab", font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, pady=(0, 10))
        
        # Create notebook for tabs
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Add first tab
        self.add_new_tab()
        
        # Tab management buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=2, column=0, pady=(10, 0), sticky=(tk.W, tk.E))
        
        add_tab_button = ttk.Button(button_frame, text="Add New Tab", command=self.add_new_tab, width=15)
        add_tab_button.pack(side=tk.LEFT, padx=(0, 10))
        
        remove_tab_button = ttk.Button(button_frame, text="Remove Current Tab", command=self.remove_current_tab, width=15)
        remove_tab_button.pack(side=tk.LEFT, padx=(0, 10))
        
        # Save/Load buttons
        save_button = ttk.Button(button_frame, text="Save All Tabs", command=self.save_all_tabs, width=15)
        save_button.pack(side=tk.LEFT, padx=(0, 10))
        
        load_button = ttk.Button(button_frame, text="Load All Tabs", command=self.load_all_tabs, width=15)
        load_button.pack(side=tk.LEFT, padx=(0, 10))
        
        # Merge all button
        merge_all_button = ttk.Button(button_frame, text="Merge All Tabs", command=self.merge_all_tabs, width=15)
        merge_all_button.pack(side=tk.LEFT)
        
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
        
    def remove_current_tab(self):
        """Remove the currently selected tab"""
        current_tab = self.notebook.select()
        if current_tab:
            tab_name = self.notebook.tab(current_tab, "text")
            
            # Don't remove if it's the last tab
            if len(self.tabs) > 1:
                self.notebook.forget(current_tab)
                del self.tabs[tab_name]
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
                
                for order, file_path, filename, file_type in tab_content.selected_files:
                    # Update file status
                    self.root.after(0, lambda f=filename, t=tab_name: self.update_tab_file_status(f, t, 'Converting...'))
                    
                    pdf_path = self.convert_to_pdf(file_path, file_type)
                    if pdf_path:
                        merger.append(pdf_path)
                        temp_files.append(pdf_path)
                        self.root.after(0, lambda f=filename, t=tab_name: self.update_tab_file_status(f, t, 'Converted'))
                    else:
                        self.root.after(0, lambda f=filename, t=tab_name: self.update_tab_file_status(f, t, 'Failed'))
                        
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
                
    def convert_to_pdf(self, file_path, file_type):
        file_ext = Path(file_path).suffix.lower()
        
        if file_ext == '.pdf':
            return file_path
        elif file_ext in ['.xlsx', '.xls', '.xlsm']:
            return self.convert_excel_to_pdf(file_path)
        elif file_ext in ['.doc', '.docx']:
            return self.convert_word_to_pdf(file_path)
        else:
            return None
            
    def convert_excel_to_pdf(self, file_path):
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
            
            # Use win32com to convert Excel to PDF
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            
            workbook = excel.Workbooks.Open(normalized_path)
            workbook.ExportAsFixedFormat(0, temp_pdf.name)  # 0 = PDF format
            workbook.Close()
            excel.Quit()
            
            return temp_pdf.name
            
        except Exception as e:
            print(f"Error converting Excel file '{file_path}': {e}")
            return None
            
    def convert_word_to_pdf(self, file_path):
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
            
            # Use win32com to convert Word to PDF
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            word.DisplayAlerts = False
            
            doc = word.Documents.Open(normalized_path)
            doc.SaveAs(temp_pdf.name, FileFormat=17)  # 17 = PDF format
            doc.Close()
            word.Quit()
            
            return temp_pdf.name
            
        except Exception as e:
            print(f"Error converting Word file '{file_path}': {e}")
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
    root.mainloop()

if __name__ == "__main__":
    main()
