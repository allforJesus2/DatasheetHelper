import tkinter as tk
from tkinter import ttk

class ModelCodeGenerator:
    def __init__(self, root):
        self.root = root
        self.root.title("MT3809G Model Code Generator")
        self.root.geometry("650x700")
        self.root.minsize(600, 500)
        
        # Configure style for modern look
        style = ttk.Style()
        style.theme_use('clam')
        
        # Create scrollable canvas
        self.canvas = tk.Canvas(root)
        self.scrollbar = ttk.Scrollbar(root, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)
        
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        
        # Pack scrollable area
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")
        
        # Enable mouse wheel scrolling
        self.canvas.bind_all("<MouseWheel>", self.on_mousewheel)
        
        # Selections storage
        self.selections = {}
        
        # Build UI
        self.create_widgets()
        
    def on_mousewheel(self, event):
        """Handle mouse wheel scrolling"""
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        
    def create_widgets(self):
        """Create all UI widgets"""
        # Title
        title = ttk.Label(self.scrollable_frame, text="MT3809G Model Code Configurator", 
                         font=('Segoe UI', 16, 'bold'))
        title.pack(pady=(10, 5))
        
        subtitle = ttk.Label(self.scrollable_frame, 
                            text="Configure each position to generate the complete model code",
                            font=('Segoe UI', 10, 'italic'))
        subtitle.pack(pady=(0, 10))
        
        # Separator
        ttk.Separator(self.scrollable_frame, orient='horizontal').pack(fill='x', pady=10)
        
        # Options frame
        options_frame = ttk.Frame(self.scrollable_frame)
        options_frame.pack(fill='x', padx=15)
        
        # Model code position definitions (Based on Data Sheet MT3809G)
        # Format: Position: {label, type, value/options}
        self.positions = {
            'I-IV': {'label': 'Model Series', 'type': 'fixed', 'value': '3809'},
            'V': {'label': 'Type', 'type': 'fixed', 'value': 'G'},
            'VI': {
                'label': 'Connection / Flange',
                'options': [
                    ('A', 'Standard Connection'),
                    ('B', 'Alternative Connection'),
                    ('Q', 'ANSI 900/1500LBS RF Flange'),
                    ('R', 'ANSI 900/1500LBS RTJ Flange'),
                    ('S', 'ANSI 2500LBS RTJ Flange'),
                    ('T', 'Threaded Connection'),
                    ('U', 'Union Connection'),
                    ('V', 'Victaulic Connection'),
                    ('W', 'Wafer Connection'),
                ]
            },
            'VII': {
                'label': 'Material',
                'options': [
                    ('A', '316L Stainless Steel'),
                    ('B', '316L SS Dual Certified (Titanium Float)'),
                    ('C', 'Hastelloy C-276 (CODE 5*)'),
                    ('D', 'Inconel 625 (CODE 5*)'),
                    ('E', 'Titanium Grade II (CODE 5*)'),
                    ('F', 'Special Alloy'),
                ]
            },
            'VIII-IX': {
                'label': 'Size Code',
                'options': [
                    ('01', 'Size 01 (1/2" NPT)'),
                    ('02', 'Size 02 (3/4" NPT)'),
                    ('03', 'Size 03 (1" NPT)'),
                    ('04', 'Size 04 (1-1/2" Flange)'),
                    ('05', 'Size 05 (2" Flange)'),
                    ('06', 'Size 06 (3" Flange)'),
                    ('07', 'Size 07 (4" Flange)'),
                    ('08', 'Size 08 (6" Flange)'),
                    ('09', 'Size 09 (8" Flange)'),
                    ('10', 'Size 10 (10" Flange)'),
                ]
            },
            'X': {
                'label': 'Certificate',
                'options': [
                    ('A', 'No Certificate'),
                    ('B', 'Material Certificate 3.1'),
                    ('C', 'CODE 5* Certificate'),
                ]
            },
            'XI': {
                'label': 'Float Detail',
                'options': [
                    ('A', 'Standard Float Detail'),
                    ('F', 'Titanium Float (Lightweight)'),
                    ('S', 'Special Float Configuration'),
                ]
            },
            'XII': {
                'label': 'CRN Registration',
                'options': [
                    ('A', 'No CRN (Standard)'),
                    ('C', 'CRN Approved (Canada)'),
                ]
            },
            'XIII': {
                'label': 'Pressure Vessel Code',
                'options': [
                    ('A', 'Standard ASME'),
                    ('C', 'CODE 5* Pressure Vessel'),
                    ('E', 'Enhanced Pressure Vessel'),
                ]
            },
            'XIV': {
                'label': 'Inspection Certificate',
                'options': [
                    ('2', 'Declaration of Compliance (EN 10204 2.1)'),
                    ('3', 'Inspection Certificate (EN 10204 3.1)'),
                    ('P', 'PMI - Positive Material Identification'),
                    ('M', 'PAMI - Positive Alloy Material ID (Carbon)'),
                    ('N', 'NACE MR0175/ISO 15156'),
                ]
            },
            'XV': {
                'label': 'NACE Compliance',
                'options': [
                    ('A', 'No NACE Compliance'),
                    ('E', 'NACE MR0175/103 Only'),
                    ('N', 'NACE + PMI Combined'),
                ]
            },
            'XVI': {
                'label': 'Cleaning Standard',
                'options': [
                    ('0', 'Standard Factory Clean'),
                    ('4', 'Commercial Clean (Standard)'),
                    ('6', 'Oxygen Service Clean'),
                    ('7', 'High Purity Clean'),
                ]
            },
            'XVII': {
                'label': 'Special Feature 1',
                'options': [
                    ('A', 'No Special Feature'),
                    ('C', 'Special Feature C (Extended)'),
                    ('D', 'Special Feature D (Compact)'),
                ]
            },
            'XVIII': {
                'label': 'Special Feature 2',
                'options': [
                    ('0', 'No Secondary Feature'),
                    ('1', 'Secondary Feature 1'),
                    ('2', 'Secondary Feature 2'),
                ]
            },
            'XIX': {
                'label': 'Special Feature 3',
                'options': [
                    ('A', 'No Tertiary Feature'),
                    ('B', 'Tertiary Feature B'),
                    ('C', 'Tertiary Feature C'),
                ]
            },
            'XX': {
                'label': 'Final Option',
                'options': [
                    ('A', 'Final Option A (Standard)'),
                    ('B', 'Final Option B (Enhanced)'),
                    ('C', 'Final Option C (Premium)'),
                ]
            },
        }
        
        # Create UI elements for each position
        for pos, data in self.positions.items():
            frame = ttk.Frame(options_frame)
            frame.pack(fill='x', pady=4)
            
            # Position label
            pos_label = ttk.Label(frame, text=f"Pos {pos}:", 
                                 font=('Segoe UI', 9, 'bold'), width=10, anchor='w')
            pos_label.pack(side='left')
            
            # Description label
            desc_label = ttk.Label(frame, text=data['label'], 
                                  font=('Segoe UI', 9), width=20, anchor='w')
            desc_label.pack(side='left')
            
            if data.get('type') == 'fixed':
                # Fixed value display
                value_label = ttk.Label(frame, text=f"Fixed: {data['value']}", 
                                       font=('Courier', 10, 'bold'), 
                                       foreground='darkblue')
                value_label.pack(side='left', padx=10)
                
                self.selections[pos] = tk.StringVar(value=data['value'])
            else:
                # Dropdown selection
                self.selections[pos] = tk.StringVar()
                
                # Set default to first option
                first_option = data['options'][0]
                self.selections[pos].set(f"{first_option[0]} - {first_option[1]}")
                
                # Format options for display
                formatted_options = [f"{code} - {desc}" for code, desc in data['options']]
                
                dropdown = ttk.Combobox(frame, textvariable=self.selections[pos],
                                       values=formatted_options, state='readonly',
                                       width=45, font=('Segoe UI', 9))
                dropdown.pack(side='left', padx=10)
                
                # Bind selection change
                dropdown.bind('<<ComboboxSelected>>', self.update_model_code)
            
            # Add separator line
            ttk.Separator(options_frame, orient='horizontal').pack(fill='x', pady=2)
        
        # Model code display section
        display_frame = ttk.Frame(self.scrollable_frame)
        display_frame.pack(fill='x', padx=15, pady=20)
        
        ttk.Separator(display_frame, orient='horizontal').pack(fill='x', pady=10)
        
        model_label = ttk.Label(display_frame, text="Generated Model Code", 
                               font=('Segoe UI', 14, 'bold'))
        model_label.pack(pady=10)
        
        # Model code entry
        self.model_code_var = tk.StringVar()
        model_entry = ttk.Entry(display_frame, textvariable=self.model_code_var,
                               font=('Courier New', 16, 'bold'), state='readonly',
                               justify='center', width=25)
        model_entry.pack(pady=10)
        
        # Button frame
        button_frame = ttk.Frame(display_frame)
        button_frame.pack(pady=10)
        
        # Copy button
        copy_btn = ttk.Button(button_frame, text="📋 Copy to Clipboard", 
                             command=self.copy_to_clipboard, width=20)
        copy_btn.grid(row=0, column=0, padx=5)
        
        # Reset button
        reset_btn = ttk.Button(button_frame, text="🔄 Reset to Defaults", 
                              command=self.reset_defaults, width=20)
        reset_btn.grid(row=0, column=1, padx=5)
        
        # Status label
        self.status_var = tk.StringVar()
        self.status_label = ttk.Label(display_frame, textvariable=self.status_var,
                                     font=('Segoe UI', 10))
        self.status_label.pack(pady=5)
        
        # Initialize model code
        self.update_model_code()
        
    def update_model_code(self, event=None):
        """Generate model code from all selections"""
        model_code = ""
        
        for pos, data in self.positions.items():
            if data.get('type') == 'fixed':
                model_code += data['value']
            else:
                # Extract code from "Code - Description" format
                selected = self.selections[pos].get()
                code_part = selected.split(' - ')[0]
                
                # Handle two-character position (VIII-IX)
                if pos == 'VIII-IX':
                    model_code += code_part.ljust(2, '0')[:2]
                else:
                    model_code += code_part
        
        self.model_code_var.set(model_code)
        
    def copy_to_clipboard(self):
        """Copy model code to clipboard with feedback"""
        try:
            model_code = self.model_code_var.get()
            self.root.clipboard_clear()
            self.root.clipboard_append(model_code)
            self.root.update()  # Required for clipboard on some platforms
            
            self.status_var.set("✓ Model code copied to clipboard!")
            self.status_label.config(foreground='green')
            self.root.after(2500, lambda: self.status_var.set(""))
        except Exception as e:
            self.status_var.set("✗ Failed to copy to clipboard")
            self.status_label.config(foreground='red')
            self.root.after(2500, lambda: self.status_var.set(""))
            
    def reset_defaults(self):
        """Reset all dropdowns to their default values"""
        for pos, data in self.positions.items():
            if data.get('type') != 'fixed' and 'options' in data:
                default_option = data['options'][0]
                self.selections[pos].set(f"{default_option[0]} - {default_option[1]}")
        
        self.update_model_code()
        self.status_var.set("✓ Reset to default values")
        self.status_label.config(foreground='blue')
        self.root.after(2500, lambda: self.status_var.set(""))


if __name__ == "__main__":
    root = tk.Tk()
    app = ModelCodeGenerator(root)
    root.mainloop()