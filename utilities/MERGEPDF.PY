import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from PyPDF2 import PdfMerger, PdfReader
import threading

class PDFMergerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Merger Tool")
        self.root.geometry("800x600")
        self.root.configure(bg='#f0f0f0')
        
        # Variables
        self.pdf_files = []
        self.output_filename = tk.StringVar(value="merged.pdf")
        
        # Configure style
        self.setup_styles()
        
        # Create GUI
        self.create_widgets()
        
    def setup_styles(self):
        """Configure ttk styles for a modern look"""
        style = ttk.Style()
        style.theme_use('clam')
        
        # Configure button styles
        style.configure('Action.TButton', 
                       padding=10, 
                       font=('Segoe UI', 10, 'bold'))
        
        style.configure('File.TButton', 
                       padding=8, 
                       font=('Segoe UI', 9))
        
        # Configure frame styles
        style.configure('Card.TFrame', 
                       background='white', 
                       relief='raised', 
                       borderwidth=1)
    
    def create_widgets(self):
        """Create and arrange all GUI widgets"""
        # Main container
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Title
        title_label = ttk.Label(main_frame, 
                               text="PDF Merger Tool", 
                               font=('Segoe UI', 20, 'bold'))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # File selection section
        self.create_file_selection_section(main_frame)
        
        # File list section
        self.create_file_list_section(main_frame)
        
        # Output section
        self.create_output_section(main_frame)
        
        # Action buttons
        self.create_action_buttons(main_frame)
        
        # Status section
        self.create_status_section(main_frame)
    
    def create_file_selection_section(self, parent):
        """Create the file selection section"""
        # File selection frame
        file_frame = ttk.LabelFrame(parent, text="File Selection", padding="15")
        file_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 15))
        file_frame.columnconfigure(1, weight=1)
        
        # Add files button
        add_files_btn = ttk.Button(file_frame, 
                                  text="Add PDF Files", 
                                  style='Action.TButton',
                                  command=self.add_files)
        add_files_btn.grid(row=0, column=0, padx=(0, 10))
        
        # Add folder button
        add_folder_btn = ttk.Button(file_frame, 
                                   text="Add Folder", 
                                   style='Action.TButton',
                                   command=self.add_folder)
        add_folder_btn.grid(row=0, column=1, padx=(0, 10))
        
        # Clear all button
        clear_btn = ttk.Button(file_frame, 
                              text="Clear All", 
                              style='File.TButton',
                              command=self.clear_files)
        clear_btn.grid(row=0, column=2)
        
        # Instructions
        instructions = ttk.Label(file_frame, 
                               text="Select individual PDF files or add all PDFs from a folder",
                               font=('Segoe UI', 9),
                               foreground='gray')
        instructions.grid(row=1, column=0, columnspan=3, pady=(10, 0), sticky=tk.W)
    
    def create_file_list_section(self, parent):
        """Create the file list display section"""
        # File list frame
        list_frame = ttk.LabelFrame(parent, text="Selected Files", padding="15")
        list_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 15))
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)
        
        # Create treeview for file list
        columns = ('#', 'Filename', 'Path', 'Size')
        self.file_tree = ttk.Treeview(list_frame, columns=columns, show='headings', height=8)
        
        # Configure columns
        self.file_tree.heading('#', text='#')
        self.file_tree.heading('Filename', text='Filename')
        self.file_tree.heading('Path', text='Path')
        self.file_tree.heading('Size', text='Size')
        
        self.file_tree.column('#', width=50, minwidth=50)
        self.file_tree.column('Filename', width=200, minwidth=150)
        self.file_tree.column('Path', width=300, minwidth=200)
        self.file_tree.column('Size', width=100, minwidth=80)
        
        # Scrollbars
        tree_scroll_y = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.file_tree.yview)
        tree_scroll_x = ttk.Scrollbar(list_frame, orient=tk.HORIZONTAL, command=self.file_tree.xview)
        self.file_tree.configure(yscrollcommand=tree_scroll_y.set, xscrollcommand=tree_scroll_x.set)
        
        # Grid layout
        self.file_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        tree_scroll_y.grid(row=0, column=1, sticky=(tk.N, tk.S))
        tree_scroll_x.grid(row=1, column=0, sticky=(tk.W, tk.E))
        
        # File list buttons
        btn_frame = ttk.Frame(list_frame)
        btn_frame.grid(row=2, column=0, columnspan=2, pady=(10, 0), sticky=tk.W)
        
        remove_btn = ttk.Button(btn_frame, 
                               text="Remove Selected", 
                               style='File.TButton',
                               command=self.remove_selected)
        remove_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        move_up_btn = ttk.Button(btn_frame, 
                                text="Move Up", 
                                style='File.TButton',
                                command=self.move_up)
        move_up_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        move_down_btn = ttk.Button(btn_frame, 
                                  text="Move Down", 
                                  style='File.TButton',
                                  command=self.move_down)
        move_down_btn.pack(side=tk.LEFT)
    
    def create_output_section(self, parent):
        """Create the output filename section"""
        # Output frame
        output_frame = ttk.LabelFrame(parent, text="Output Settings", padding="15")
        output_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 15))
        output_frame.columnconfigure(1, weight=1)
        
        # Output filename label and entry
        ttk.Label(output_frame, text="Output Filename:").grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        
        output_entry = ttk.Entry(output_frame, textvariable=self.output_filename, font=('Segoe UI', 10))
        output_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 10))
        
        # Browse button for output location
        browse_btn = ttk.Button(output_frame, 
                               text="Browse", 
                               style='File.TButton',
                               command=self.browse_output)
        browse_btn.grid(row=0, column=2)
        
        # Ensure .pdf extension
        ttk.Label(output_frame, 
                 text="(Will automatically add .pdf extension if missing)",
                 font=('Segoe UI', 8),
                 foreground='gray').grid(row=1, column=1, sticky=tk.W, pady=(5, 0))
    
    def create_action_buttons(self, parent):
        """Create the main action buttons"""
        # Action frame
        action_frame = ttk.Frame(parent)
        action_frame.grid(row=4, column=0, columnspan=3, pady=(0, 15))
        
        # Merge button
        self.merge_btn = ttk.Button(action_frame, 
                                   text="Merge PDFs", 
                                   style='Action.TButton',
                                   command=self.merge_pdfs)
        self.merge_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # Progress bar
        self.progress = ttk.Progressbar(action_frame, mode='indeterminate')
        self.progress.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        
        # Status label
        self.status_label = ttk.Label(action_frame, text="Ready", font=('Segoe UI', 9))
        self.status_label.pack(side=tk.RIGHT)
    
    def create_status_section(self, parent):
        """Create the status/log section"""
        # Status frame
        status_frame = ttk.LabelFrame(parent, text="Status Log", padding="15")
        status_frame.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S))
        status_frame.columnconfigure(0, weight=1)
        status_frame.rowconfigure(0, weight=1)
        
        # Status text area
        self.status_text = scrolledtext.ScrolledText(status_frame, 
                                                    height=8, 
                                                    font=('Consolas', 9),
                                                    wrap=tk.WORD)
        self.status_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure parent grid weights
        parent.rowconfigure(5, weight=1)
    
    def add_files(self):
        """Add individual PDF files"""
        files = filedialog.askopenfilenames(
            title="Select PDF Files",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        
        if files:
            for file in files:
                if file not in self.pdf_files:
                    self.pdf_files.append(file)
            self.update_file_list()
            self.log_message(f"Added {len(files)} file(s)")
    
    def add_folder(self):
        """Add all PDF files from a folder"""
        folder = filedialog.askdirectory(title="Select Folder Containing PDFs")
        
        if folder:
            pdf_files = []
            for filename in sorted(os.listdir(folder)):
                if filename.lower().endswith('.pdf'):
                    file_path = os.path.join(folder, filename)
                    if file_path not in self.pdf_files:
                        pdf_files.append(file_path)
            
            if pdf_files:
                self.pdf_files.extend(pdf_files)
                self.update_file_list()
                self.log_message(f"Added {len(pdf_files)} PDF file(s) from folder")
            else:
                messagebox.showwarning("No PDFs Found", "No PDF files found in the selected folder.")
    
    def clear_files(self):
        """Clear all selected files"""
        if self.pdf_files:
            self.pdf_files.clear()
            self.update_file_list()
            self.log_message("Cleared all files")
    
    def remove_selected(self):
        """Remove selected file from list"""
        selection = self.file_tree.selection()
        if selection:
            for item in selection:
                index = int(self.file_tree.item(item)['values'][0]) - 1
                if 0 <= index < len(self.pdf_files):
                    removed_file = self.pdf_files.pop(index)
                    self.log_message(f"Removed: {os.path.basename(removed_file)}")
            self.update_file_list()
    
    def move_up(self):
        """Move selected file up in the list"""
        selection = self.file_tree.selection()
        if selection:
            index = int(self.file_tree.item(selection[0])['values'][0]) - 1
            if index > 0:
                self.pdf_files[index], self.pdf_files[index-1] = self.pdf_files[index-1], self.pdf_files[index]
                self.update_file_list()
                self.file_tree.selection_set(self.file_tree.get_children()[index-1])
    
    def move_down(self):
        """Move selected file down in the list"""
        selection = self.file_tree.selection()
        if selection:
            index = int(self.file_tree.item(selection[0])['values'][0]) - 1
            if index < len(self.pdf_files) - 1:
                self.pdf_files[index], self.pdf_files[index+1] = self.pdf_files[index+1], self.pdf_files[index]
                self.update_file_list()
                self.file_tree.selection_set(self.file_tree.get_children()[index+1])
    
    def browse_output(self):
        """Browse for output file location"""
        filename = filedialog.asksaveasfilename(
            title="Save Merged PDF As",
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if filename:
            self.output_filename.set(filename)
    
    def update_file_list(self):
        """Update the file list display"""
        # Clear existing items
        for item in self.file_tree.get_children():
            self.file_tree.delete(item)
        
        # Add files to treeview
        for i, file_path in enumerate(self.pdf_files, 1):
            try:
                file_size = os.path.getsize(file_path)
                size_str = self.format_file_size(file_size)
                self.file_tree.insert('', 'end', values=(
                    i,
                    os.path.basename(file_path),
                    file_path,
                    size_str
                ))
            except OSError:
                self.file_tree.insert('', 'end', values=(
                    i,
                    os.path.basename(file_path),
                    file_path,
                    "Error"
                ))
        
        # Update merge button state
        self.merge_btn.config(state='normal' if len(self.pdf_files) >= 2 else 'disabled')
    
    def format_file_size(self, size_bytes):
        """Format file size in human readable format"""
        if size_bytes == 0:
            return "0 B"
        
        size_names = ["B", "KB", "MB", "GB"]
        i = 0
        while size_bytes >= 1024 and i < len(size_names) - 1:
            size_bytes /= 1024.0
            i += 1
        
        return f"{size_bytes:.1f} {size_names[i]}"
    
    def log_message(self, message):
        """Add message to status log"""
        self.status_text.insert(tk.END, f"{message}\n")
        self.status_text.see(tk.END)
        self.root.update_idletasks()
    
    def merge_pdfs(self):
        """Merge the selected PDF files"""
        if len(self.pdf_files) < 2:
            messagebox.showwarning("Insufficient Files", "Please select at least 2 PDF files to merge.")
            return
        
        # Get output filename
        output_name = self.output_filename.get().strip()
        if not output_name:
            messagebox.showerror("Error", "Please specify an output filename.")
            return
        
        if not output_name.endswith('.pdf'):
            output_name += '.pdf'
        
        # Start merge in separate thread
        self.merge_btn.config(state='disabled')
        self.progress.start()
        self.status_label.config(text="Merging...")
        
        thread = threading.Thread(target=self._merge_pdfs_thread, args=(output_name,))
        thread.daemon = True
        thread.start()
    
    def _merge_pdfs_thread(self, output_name):
        """Merge PDFs in a separate thread"""
        try:
            merger = PdfMerger()
            successful_files = 0
            
            for pdf in self.pdf_files:
                if os.path.exists(pdf):
                    self.log_message(f"Adding: {os.path.basename(pdf)}")
                    merger.append(pdf)
                    successful_files += 1
                else:
                    self.log_message(f"Warning: File not found - {pdf}")
            
            if successful_files > 0:
                # Write the merged PDF
                with open(output_name, 'wb') as output_pdf:
                    merger.write(output_pdf)
                
                self.log_message(f"\nSuccessfully merged {successful_files} PDFs into '{output_name}'")
                messagebox.showinfo("Success", f"PDFs merged successfully!\nOutput: {output_name}")
            else:
                self.log_message("No valid PDF files to merge!")
                messagebox.showerror("Error", "No valid PDF files found to merge.")
            
        except Exception as e:
            error_msg = f"Error merging PDFs: {e}"
            self.log_message(error_msg)
            messagebox.showerror("Error", error_msg)
        
        finally:
            merger.close()
            # Update GUI in main thread
            self.root.after(0, self._merge_complete)
    
    def _merge_complete(self):
        """Called when merge operation completes"""
        self.progress.stop()
        self.merge_btn.config(state='normal')
        self.status_label.config(text="Ready")

def main():
    root = tk.Tk()
    app = PDFMergerGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
