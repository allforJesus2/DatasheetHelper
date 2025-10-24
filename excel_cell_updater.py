import tkinter as tk
from tkinter import messagebox, filedialog, ttk
import xlwings as xw
import difflib
import re
import os


class ExcelCellUpdater:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Cell Updater")
        self.root.geometry("600x500")
        
        # Excel workbook and selection tracking
        self.wb = None
        self.current_selection = None
        
        # Cell references
        self.old_cell = None
        self.new_value_cell = None
        self.destination_cell = None
        
        # File path for remembering last opened file
        self.last_file_path = self.load_last_file_path()
        
        # Create GUI elements
        self.create_widgets()
        
        # Auto-load last file if it exists
        self.auto_load_last_file()
        
        # Start monitoring Excel selection
        self.monitor_excel_selection()
    
    def load_last_file_path(self):
        """Load the last opened file path from a text file"""
        try:
            config_file = "last_excel_file.txt"
            if os.path.exists(config_file):
                with open(config_file, 'r') as f:
                    file_path = f.read().strip()
                    if os.path.exists(file_path):
                        return file_path
        except Exception as e:
            print(f"Error loading last file path: {e}")
        return None
    
    def save_last_file_path(self, file_path):
        """Save the last opened file path to a text file"""
        try:
            config_file = "last_excel_file.txt"
            with open(config_file, 'w') as f:
                f.write(file_path)
        except Exception as e:
            print(f"Error saving last file path: {e}")
    
    def auto_load_last_file(self):
        """Automatically load the last opened file if it exists"""
        if self.last_file_path:
            try:
                self.wb = xw.Book(self.last_file_path)
                self.file_label.config(text=os.path.basename(self.last_file_path), fg="black")
                self.status_label.config(text="Last file loaded automatically", fg="green")
                
                # Enable buttons
                self.set_old_btn.config(state=tk.NORMAL)
                self.set_new_btn.config(state=tk.NORMAL)
                self.set_dest_btn.config(state=tk.NORMAL)
                
            except Exception as e:
                print(f"Error auto-loading last file: {e}")
                self.status_label.config(text="Failed to auto-load last file", fg="orange")
    
    def create_widgets(self):
        # File selection frame
        file_frame = tk.Frame(self.root)
        file_frame.pack(fill=tk.X, padx=10, pady=5)
        
        tk.Label(file_frame, text="Excel File:").pack(side=tk.LEFT)
        self.file_label = tk.Label(file_frame, text="No file selected", fg="gray")
        self.file_label.pack(side=tk.LEFT, padx=5)
        
        tk.Button(file_frame, text="Browse", command=self.browse_file).pack(side=tk.RIGHT)
        
        # Current selection frame
        selection_frame = tk.Frame(self.root)
        selection_frame.pack(fill=tk.X, padx=10, pady=5)
        
        tk.Label(selection_frame, text="Current Selection:").pack(side=tk.LEFT)
        self.selection_label = tk.Label(selection_frame, text="No cell selected", fg="gray")
        self.selection_label.pack(side=tk.LEFT, padx=5)
        
        # Old cell frame
        old_cell_frame = tk.Frame(self.root)
        old_cell_frame.pack(fill=tk.X, padx=10, pady=5)
        
        tk.Label(old_cell_frame, text="Old Cell:").pack(side=tk.LEFT)
        self.old_cell_label = tk.Label(old_cell_frame, text="Not set", fg="red", bg="lightgray", relief=tk.SUNKEN, padx=5, pady=2)
        self.old_cell_label.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(5, 10))
        
        self.set_old_btn = tk.Button(old_cell_frame, text="Set Old Cell", command=self.set_old_cell, state=tk.DISABLED)
        self.set_old_btn.pack(side=tk.RIGHT)
        
        # New value cell frame
        new_cell_frame = tk.Frame(self.root)
        new_cell_frame.pack(fill=tk.X, padx=10, pady=5)
        
        tk.Label(new_cell_frame, text="New Value Cell:").pack(side=tk.LEFT)
        self.new_cell_label = tk.Label(new_cell_frame, text="Not set", fg="red", bg="lightgray", relief=tk.SUNKEN, padx=5, pady=2)
        self.new_cell_label.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(5, 10))
        
        self.set_new_btn = tk.Button(new_cell_frame, text="Set New Cell", command=self.set_new_cell, state=tk.DISABLED)
        self.set_new_btn.pack(side=tk.RIGHT)
        
        # Destination cell frame
        dest_frame = tk.Frame(self.root)
        dest_frame.pack(fill=tk.X, padx=10, pady=5)
        
        tk.Label(dest_frame, text="Destination Cell:").pack(side=tk.LEFT)
        self.dest_label = tk.Label(dest_frame, text="Not set", fg="red", bg="lightgray", relief=tk.SUNKEN, padx=5, pady=2)
        self.dest_label.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(5, 10))
        
        self.set_dest_btn = tk.Button(dest_frame, text="Set Destination Cell", command=self.set_destination_cell, state=tk.DISABLED)
        self.set_dest_btn.pack(side=tk.RIGHT)
        
        # Values display frame
        values_frame = tk.Frame(self.root)
        values_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # Old value display
        old_value_frame = tk.Frame(values_frame)
        old_value_frame.pack(fill=tk.X, pady=2)
        
        tk.Label(old_value_frame, text="Old Value:").pack(side=tk.LEFT)
        self.old_value_label = tk.Label(old_value_frame, text="", fg="blue", bg="lightyellow", relief=tk.SUNKEN, padx=5, pady=2)
        self.old_value_label.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(5, 10))
        
        # New value display
        new_value_frame = tk.Frame(values_frame)
        new_value_frame.pack(fill=tk.X, pady=2)
        
        tk.Label(new_value_frame, text="New Value:").pack(side=tk.LEFT)
        self.new_value_label = tk.Label(new_value_frame, text="", fg="green", bg="lightyellow", relief=tk.SUNKEN, padx=5, pady=2)
        self.new_value_label.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(5, 10))
        
        # Preview frame
        preview_frame = tk.LabelFrame(self.root, text="Preview - How the text will look after update", padx=10, pady=5)
        preview_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # Preview text widget with scrollbar
        preview_text_frame = tk.Frame(preview_frame)
        preview_text_frame.pack(fill=tk.BOTH, expand=True)
        
        self.preview_text = tk.Text(preview_text_frame, height=6, wrap=tk.WORD, state=tk.DISABLED)
        preview_scrollbar = tk.Scrollbar(preview_text_frame, orient=tk.VERTICAL, command=self.preview_text.yview)
        self.preview_text.configure(yscrollcommand=preview_scrollbar.set)
        
        self.preview_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        preview_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Preview button
        preview_button_frame = tk.Frame(preview_frame)
        preview_button_frame.pack(fill=tk.X, pady=5)
        
        self.preview_button = tk.Button(preview_button_frame, text="Generate Preview", command=self.generate_preview, state=tk.DISABLED, bg="lightgreen")
        self.preview_button.pack(side=tk.RIGHT)
        
        # Options frame
        options_frame = tk.LabelFrame(self.root, text="Options", padx=10, pady=5)
        options_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # Only update when different checkbox
        self.only_when_different_var = tk.BooleanVar(value=True)
        self.only_when_different_checkbox = tk.Checkbutton(
            options_frame, 
            text="Only update destination cell when old and new values are different", 
            variable=self.only_when_different_var,
            fg="blue"
        )
        self.only_when_different_checkbox.pack(anchor=tk.W)
        
        # Update button frame
        button_frame = tk.Frame(self.root)
        button_frame.pack(fill=tk.X, padx=10, pady=10)
        
        self.update_button = tk.Button(button_frame, text="Update Destination Cell", command=self.update_destination_cell, state=tk.DISABLED, bg="lightblue", fg="black")
        self.update_button.pack(side=tk.RIGHT)
        
        # Batch update button
        self.batch_button = tk.Button(button_frame, text="Batch Update Down Columns", command=self.batch_update_down_columns, state=tk.DISABLED, bg="orange", fg="black")
        self.batch_button.pack(side=tk.RIGHT, padx=(0, 10))
        
        # Status frame
        status_frame = tk.Frame(self.root)
        status_frame.pack(fill=tk.X, padx=10, pady=5)
        
        self.status_label = tk.Label(status_frame, text="Ready - Select cells and press buttons to set them", fg="blue")
        self.status_label.pack(side=tk.LEFT)
    
    def generate_preview(self):
        """Generate a preview of how the text will look with color coding"""
        if not self.old_cell or not self.new_value_cell:
            messagebox.showwarning("Warning", "Please set both old cell and new value cell first.")
            return
        
        try:
            # Get the old and new values
            old_cell_range = self.wb.sheets.active.range(self.old_cell)
            new_cell_range = self.wb.sheets.active.range(self.new_value_cell)
            
            old_value = self.get_cell_value_with_merge_handling(old_cell_range)
            new_value = self.get_cell_value_with_merge_handling(new_cell_range)
            
            old_text = str(old_value) if old_value is not None else ""
            new_text = str(new_value) if new_value is not None else ""
            
            # Clear previous preview
            self.preview_text.config(state=tk.NORMAL)
            self.preview_text.delete(1.0, tk.END)
            
            # Add header
            self.preview_text.insert(tk.END, "Preview of how the destination cell will look:\n\n")
            
            # Generate colored preview
            self.create_text_preview(old_text, new_text)
            
            self.preview_text.config(state=tk.DISABLED)
            self.status_label.config(text="Preview generated successfully", fg="green")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate preview: {str(e)}")
            self.status_label.config(text="Failed to generate preview", fg="red")
    
    def create_text_preview(self, old_text, new_text):
        """Create a text preview showing character-level differences"""
        try:
            # Use difflib to find differences
            matcher = difflib.SequenceMatcher(None, old_text, new_text)
            
            # Configure text widget for colored text
            self.preview_text.tag_configure("normal", foreground="black")
            self.preview_text.tag_configure("changed", foreground="red", background="lightpink")
            self.preview_text.tag_configure("added", foreground="red", background="lightpink")
            
            char_index = 0
            
            for tag, i1, i2, j1, j2 in matcher.get_opcodes():
                if tag == 'equal':
                    # Same text - normal color
                    text_segment = new_text[j1:j2]
                    self.preview_text.insert(tk.END, text_segment, "normal")
                    char_index += (j2 - j1)
                elif tag == 'replace':
                    # Replaced text - red background
                    text_segment = new_text[j1:j2]
                    self.preview_text.insert(tk.END, text_segment, "changed")
                    char_index += (j2 - j1)
                elif tag == 'delete':
                    # Deleted text - show in strikethrough style
                    text_segment = old_text[i1:i2]
                    self.preview_text.insert(tk.END, f"[DELETED: {text_segment}]", "changed")
                elif tag == 'insert':
                    # Inserted text - red background
                    text_segment = new_text[j1:j2]
                    self.preview_text.insert(tk.END, text_segment, "added")
                    char_index += (j2 - j1)
            
            # Add summary
            self.preview_text.insert(tk.END, "\n\n")
            self.preview_text.insert(tk.END, f"Summary:\n", "normal")
            self.preview_text.insert(tk.END, f"• Old text: '{old_text}'\n", "normal")
            self.preview_text.insert(tk.END, f"• New text: '{new_text}'\n", "normal")
            
            if old_text == new_text:
                self.preview_text.insert(tk.END, f"• No changes detected\n", "normal")
            else:
                self.preview_text.insert(tk.END, f"• Text will be updated with red highlighting for changes\n", "normal")
                
        except Exception as e:
            self.preview_text.insert(tk.END, f"Error generating preview: {str(e)}", "changed")
    
    def set_old_cell(self):
        """Set the current Excel selection as the old cell"""
        if self.current_selection:
            self.old_cell = self.current_selection
            self.old_cell_label.config(text=self.old_cell, fg="green", bg="lightgreen")
            
            # Get and display the old value
            try:
                cell = self.wb.sheets.active.range(self.old_cell)
                old_value = self.get_cell_value_with_merge_handling(cell)
                self.old_value_label.config(text=str(old_value) if old_value is not None else "")
            except Exception as e:
                print(f"Error getting old cell value: {e}")
                self.old_value_label.config(text="Error reading value")
            
            self.status_label.config(text=f"Old cell set to: {self.old_cell}", fg="green")
            self.check_ready_state()
    
    def set_new_cell(self):
        """Set the current Excel selection as the new value cell"""
        if self.current_selection:
            self.new_value_cell = self.current_selection
            self.new_cell_label.config(text=self.new_value_cell, fg="green", bg="lightgreen")
            
            # Get and display the new value
            try:
                cell = self.wb.sheets.active.range(self.new_value_cell)
                new_value = self.get_cell_value_with_merge_handling(cell)
                self.new_value_label.config(text=str(new_value) if new_value is not None else "")
            except Exception as e:
                print(f"Error getting new cell value: {e}")
                self.new_value_label.config(text="Error reading value")
            
            self.status_label.config(text=f"New value cell set to: {self.new_value_cell}", fg="green")
            self.check_ready_state()
    
    def set_destination_cell(self):
        """Set the current Excel selection as the destination cell"""
        if self.current_selection:
            self.destination_cell = self.current_selection
            self.dest_label.config(text=self.destination_cell, fg="green", bg="lightgreen")
            
            self.status_label.config(text=f"Destination cell set to: {self.destination_cell}", fg="green")
            self.check_ready_state()
    
    def check_ready_state(self):
        """Check if all cells are set and enable update button"""
        if self.old_cell and self.new_value_cell and self.destination_cell:
            self.update_button.config(state=tk.NORMAL)
            self.preview_button.config(state=tk.NORMAL)
            self.batch_button.config(state=tk.NORMAL)
            self.status_label.config(text="All cells set! Generate preview or click 'Update Destination Cell' to apply changes", fg="green")
        else:
            self.update_button.config(state=tk.DISABLED)
            self.preview_button.config(state=tk.DISABLED)
            self.batch_button.config(state=tk.DISABLED)
    
    def batch_update_down_columns(self):
        """Batch update down the columns for all rows in the current table"""
        if not self.wb or not self.old_cell or not self.new_value_cell or not self.destination_cell:
            messagebox.showwarning("Warning", "Please set all cells first.")
            return
        
        try:
            # Get the active sheet
            sheet = self.wb.sheets.active
            
            # Parse cell addresses to get column letters and row numbers
            old_col, old_row = self.parse_cell_address(self.old_cell)
            new_col, new_row = self.parse_cell_address(self.new_value_cell)
            dest_col, dest_row = self.parse_cell_address(self.destination_cell)
            
            # Ask user for start and end rows
            start_row = self.ask_for_row_number("Enter start row number:", old_row)
            if start_row is None:
                return
                
            end_row = self.ask_for_row_number("Enter end row number:", start_row + 10)
            if end_row is None:
                return
            
            if start_row > end_row:
                messagebox.showwarning("Warning", "Start row must be less than or equal to end row.")
                return
            
            # Confirm the operation
            result = messagebox.askyesno("Batch Update", 
                f"This will update rows {start_row} to {end_row}.\n"
                f"Old cell column: {old_col}\n"
                f"New cell column: {new_col}\n"
                f"Destination column: {dest_col}\n\n"
                f"Continue?")
            
            if not result:
                return
            
            # Perform batch update
            updated_count = 0
            for row in range(start_row, end_row + 1):
                try:
                    # Construct cell addresses for current row
                    current_old_cell = f"{old_col}{row}"
                    current_new_cell = f"{new_col}{row}"
                    current_dest_cell = f"{dest_col}{row}"
                    
                    # Get the old value from the old cell
                    old_value_cell = sheet.range(current_old_cell)
                    old_value = self.get_cell_value_with_merge_handling(old_value_cell)
                    old_value_str = str(old_value) if old_value is not None else ""
                    
                    # Get the new value from the new value cell
                    new_value_cell = sheet.range(current_new_cell)
                    new_value = self.get_cell_value_with_merge_handling(new_value_cell)
                    
                    if new_value is not None:
                        new_value_str = str(new_value)
                        
                        # Check if values are different (if the option is enabled)
                        if self.only_when_different_var.get() and old_value_str == new_value_str:
                            # Skip this row - values are identical
                            continue
                        
                        # Get the destination cell
                        dest_range = sheet.range(current_dest_cell)
                        
                        # Update the destination cell with new value, handling merged cells
                        self.update_cell_with_merge_handling(dest_range, new_value_str, old_value_str)
                        updated_count += 1
                    
                except Exception as e:
                    print(f"Error updating row {row}: {e}")
                    continue
            
            self.status_label.config(text=f"Batch update completed! Updated {updated_count} cells from row {start_row} to {end_row}", fg="green")
            messagebox.showinfo("Batch Update Complete", f"Successfully updated {updated_count} cells!")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to perform batch update: {str(e)}")
            self.status_label.config(text="Failed to perform batch update", fg="red")
    
    def ask_for_row_number(self, message, default_value):
        """Ask user for a row number using a simple dialog"""
        try:
            # Create a simple input dialog
            dialog = tk.Toplevel(self.root)
            dialog.title("Enter Row Number")
            dialog.geometry("300x150")
            dialog.transient(self.root)
            dialog.grab_set()
            
            # Center the dialog
            dialog.geometry("+%d+%d" % (self.root.winfo_rootx() + 50, self.root.winfo_rooty() + 50))
            
            # Create widgets
            tk.Label(dialog, text=message, pady=10).pack()
            
            entry_var = tk.StringVar(value=str(default_value))
            entry = tk.Entry(dialog, textvariable=entry_var, width=20)
            entry.pack(pady=5)
            entry.focus_set()
            entry.select_range(0, tk.END)
            
            result = [None]  # Use list to store result
            
            def on_ok():
                try:
                    value = int(entry_var.get())
                    if value < 1:
                        messagebox.showerror("Error", "Row number must be 1 or greater.")
                        return
                    result[0] = value
                    dialog.destroy()
                except ValueError:
                    messagebox.showerror("Error", "Please enter a valid number.")
            
            def on_cancel():
                dialog.destroy()
            
            # Buttons
            button_frame = tk.Frame(dialog)
            button_frame.pack(pady=10)
            
            tk.Button(button_frame, text="OK", command=on_ok, width=10).pack(side=tk.LEFT, padx=5)
            tk.Button(button_frame, text="Cancel", command=on_cancel, width=10).pack(side=tk.LEFT, padx=5)
            
            # Bind Enter key to OK
            entry.bind('<Return>', lambda e: on_ok())
            
            # Wait for dialog to close
            dialog.wait_window()
            
            return result[0]
            
        except Exception as e:
            print(f"Error in ask_for_row_number: {e}")
            return None
    
    def parse_cell_address(self, cell_address):
        """Parse a cell address to get column letter and row number"""
        import re
        match = re.match(r'([A-Z]+)(\d+)', cell_address.upper())
        if match:
            column = match.group(1)
            row = int(match.group(2))
            return column, row
        else:
            raise ValueError(f"Invalid cell address: {cell_address}")
    

    
    def browse_file(self):
        """Open file dialog to select Excel file"""
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls *.xlsm"), ("All files", "*.*")]
        )
        
        if file_path:
            try:
                if self.wb:
                    self.wb.close()
                
                self.wb = xw.Book(file_path)
                self.file_label.config(text=os.path.basename(file_path), fg="black")
                self.status_label.config(text="Excel file loaded successfully", fg="green")
                
                # Enable buttons
                self.set_old_btn.config(state=tk.NORMAL)
                self.set_new_btn.config(state=tk.NORMAL)
                self.set_dest_btn.config(state=tk.NORMAL)
                
                # Save the new file path
                self.save_last_file_path(file_path)
                
                # Start monitoring selection
                self.monitor_excel_selection()
                
            except Exception as e:
                messagebox.showerror("Error", f"Failed to open Excel file: {str(e)}")
                self.status_label.config(text="Failed to load Excel file", fg="red")
    
    def monitor_excel_selection(self):
        """Monitor Excel selection and update the GUI"""
        try:
            if self.wb:
                # Get current selection
                selection = self.wb.selection
                if selection:
                    # Get the address of the first cell in selection
                    address = selection.address.split(':')[0].replace('$', '')
                    if address != self.current_selection:
                        self.current_selection = address
                        self.selection_label.config(text=address, fg="black")
                        self.status_label.config(text=f"Selected cell: {address} - Click a button to set it", fg="blue")
                else:
                    if self.current_selection:
                        self.current_selection = None
                        self.selection_label.config(text="No cell selected", fg="gray")
                        self.status_label.config(text="No cell selected", fg="gray")
        
        except Exception as e:
            print(f"Error monitoring Excel selection: {e}")
        
        # Schedule next check
        self.root.after(500, self.monitor_excel_selection)
    
    def update_destination_cell(self):
        """Update the destination cell with the new value"""
        if not self.wb or not self.old_cell or not self.new_value_cell or not self.destination_cell:
            messagebox.showwarning("Warning", "Please set all cells first.")
            return
        
        try:
            # Get the old value from the old cell
            old_value_cell = self.wb.sheets.active.range(self.old_cell)
            old_value = self.get_cell_value_with_merge_handling(old_value_cell)
            old_value_str = str(old_value) if old_value is not None else ""
            
            # Get the new value from the new value cell
            new_value_cell = self.wb.sheets.active.range(self.new_value_cell)
            new_value = self.get_cell_value_with_merge_handling(new_value_cell)
            
            if new_value is None:
                messagebox.showwarning("Warning", "New value cell is empty.")
                return
            
            new_value_str = str(new_value)
            
            # Check if values are different (if the option is enabled)
            if self.only_when_different_var.get() and old_value_str == new_value_str:
                self.status_label.config(text=f"No update needed - values are identical: '{old_value_str}'", fg="blue")
                return
            
            # Get the destination cell
            dest_range = self.wb.sheets.active.range(self.destination_cell)
            
            # Update the destination cell with new value, handling merged cells
            # Use the old value from the old cell for comparison
            self.update_cell_with_merge_handling(dest_range, new_value_str, old_value_str)
            
            self.status_label.config(text=f"Successfully updated {self.destination_cell} with value from {self.new_value_cell}", fg="green")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to update destination cell: {str(e)}")
            self.status_label.config(text="Failed to update destination cell", fg="red")
    
    def get_cell_value_with_merge_handling(self, cell):
        """Get the value from a cell, handling merged cells properly"""
        try:
            # First, handle list values from xlwings selection
            if isinstance(cell.value, list):
                # This is likely a merged cell selection - get the first non-None value
                for item in cell.value:
                    if item is not None:
                        return item
                return None
            
            # Check if the cell is part of a merged range
            if hasattr(cell, 'merge_area') and cell.merge_area.address != cell.address:
                # This cell is part of a merged range but not the top-left cell
                # Get the top-left cell of the merged range
                merged_range = cell.merge_area
                top_left_address = merged_range.address.split(':')[0].replace('$', '')
                top_left_cell = self.wb.sheets.active.range(top_left_address)
                return top_left_cell.value
            else:
                # This is either a regular cell or the top-left cell of a merged range
                return cell.value
                
        except Exception as e:
            print(f"Error getting cell value with merge handling: {e}")
            return cell.value
    
    def update_cell_with_merge_handling(self, cell, new_value, old_text):
        """Update a cell with new value, handling merged cells properly"""
        try:
            # Check if the cell is part of a merged range
            if hasattr(cell, 'merge_area') and cell.merge_area.address != cell.address:
                # This cell is part of a merged range but not the top-left cell
                # Update the top-left cell of the merged range instead
                merged_range = cell.merge_area
                top_left_address = merged_range.address.split(':')[0].replace('$', '')
                top_left_cell = self.wb.sheets.active.range(top_left_address)
                
                # Update the top-left cell
                top_left_cell.value = new_value
                
                # Apply character-level color coding to the top-left cell
                self.apply_character_coloring(top_left_cell, old_text, new_value)
            else:
                # This is either a regular cell or the top-left cell of a merged range
                cell.value = new_value
                
                # Apply character-level color coding
                self.apply_character_coloring(cell, old_text, new_value)
        except Exception as e:
            print(f"Error updating cell with merge handling: {e}")
            # Fallback to regular update
            cell.value = new_value
            self.apply_character_coloring(cell, old_text, new_value)
    
    def apply_character_coloring(self, cell, old_text, new_text):
        """Apply character-level color coding to show changes"""
        try:
            # Use difflib to find differences
            matcher = difflib.SequenceMatcher(None, old_text, new_text)
            
            # Clear any existing formatting
            try:
                cell.api.Font.Color = (0, 0, 0)  # Reset to black
            except:
                pass
            
            # Apply character-level formatting
            char_index = 0
            
            for tag, i1, i2, j1, j2 in matcher.get_opcodes():
                if tag == 'equal':
                    # Same text - keep black
                    char_index += (i2 - i1)
                elif tag == 'replace':
                    # Replaced text - color the new text red
                    start_pos = char_index
                    end_pos = char_index + (j2 - j1)
                    
                    try:
                        # Color the new text red
                        if end_pos > start_pos:
                            cell.characters[start_pos:end_pos].font.color = (255, 0, 0)  # Red
                    except Exception as e:
                        print(f"Error coloring replaced text: {e}")
                    
                    char_index += (j2 - j1)
                elif tag == 'delete':
                    # Deleted text - we can't color it since it's gone, but we could mark it somehow
                    # For now, we'll just note it
                    pass
                elif tag == 'insert':
                    # Inserted text - color it red
                    start_pos = char_index
                    end_pos = char_index + (j2 - j1)
                    
                    try:
                        # Color the inserted text red
                        if end_pos > start_pos:
                            cell.characters[start_pos:end_pos].font.color = (255, 0, 0)  # Red
                    except Exception as e:
                        print(f"Error coloring inserted text: {e}")
                    
                    char_index += (j2 - j1)
            
        except Exception as e:
            print(f"Error applying character coloring: {e}")
            # Fallback: just color the entire cell red if there are changes
            try:
                if old_text != new_text:
                    cell.api.Font.Color = (255, 0, 0)  # Red
            except:
                pass


def main():
    root = tk.Tk()
    app = ExcelCellUpdater(root)
    
    # Handle window closing
    def on_closing():
        if app.wb:
            try:
                app.wb.close()
            except:
                pass
        root.destroy()
    
    root.protocol("WM_DELETE_WINDOW", on_closing)
    root.mainloop()


if __name__ == "__main__":
    main()
