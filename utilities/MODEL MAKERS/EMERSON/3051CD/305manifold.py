import tkinter as tk
from tkinter import ttk, messagebox

class Rosemount305Configurator(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Rosemount 305 Integral Manifold Configurator")
        self.geometry("900x700")
        
        # Configure Grid Weight
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        # Style configuration
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('TLabel', font=('Arial', 10))
        style.configure('Header.TLabel', font=('Arial', 12, 'bold'))
        style.configure('Result.TLabel', font=('Courier', 14, 'bold'), foreground='blue')

        # --- Main Layout with Scrollbar ---
        self.main_frame = ttk.Frame(self)
        self.main_frame.grid(row=0, column=0, sticky="nsew")
        self.main_frame.columnconfigure(0, weight=1)
        self.main_frame.rowconfigure(1, weight=1)

        # Header
        header_frame = ttk.Frame(self.main_frame, padding="10")
        header_frame.grid(row=0, column=0, sticky="ew")
        ttk.Label(header_frame, text="Rosemount 305 Integral Manifold Model Builder", style='Header.TLabel').pack()
        ttk.Label(header_frame, text="Based on Product Data Sheet 00813-0100-4733 (Rev SD, Oct 2024)", font=('Arial', 8, 'italic')).pack()

        # Canvas for Scrolling
        self.canvas = tk.Canvas(self.main_frame)
        self.scrollbar = ttk.Scrollbar(self.main_frame, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas, padding="20")

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.grid(row=1, column=0, sticky="nsew")
        self.scrollbar.grid(row=1, column=1, sticky="ns")

        # --- Data Definitions (from PDF Pages 13-16) ---
        self.init_data()

        # --- Variables ---
        self.vars = {}
        for key in self.data_structure:
            self.vars[key] = tk.StringVar()
        
        # Option vars (Checkboxes)
        self.option_vars = {}
        for cat, opts in self.options_data.items():
            for code, _ in opts:
                self.option_vars[code] = tk.BooleanVar()

        # --- UI Construction ---
        self.create_widgets()

        # --- Footer (Result) ---
        footer_frame = ttk.Frame(self.main_frame, padding="20", relief="raised")
        footer_frame.grid(row=2, column=0, sticky="ew")
        footer_frame.columnconfigure(1, weight=1)

        ttk.Label(footer_frame, text="Model Number:").grid(row=0, column=0, padx=5)
        self.result_label = ttk.Label(footer_frame, text="0305 R C 2 2 B 1 1", style='Result.TLabel')
        self.result_label.grid(row=0, column=1, sticky="w", padx=5)
        
        ttk.Button(footer_frame, text="Copy to Clipboard", command=self.copy_to_clipboard).grid(row=0, column=2, padx=5)
        
        # Initialize Default Logic
        self.vars['style'].set('C')
        self.update_logic()

    def init_data(self):
        # Ordered keys for the dropdowns
        self.data_keys = ['style', 'type', 'material', 'conn', 'packing', 'seat']
        self.data_structure = {
            'style': [
                ('C', 'Coplanar'),
                ('T', 'Traditional'),
                ('M', 'Traditional (DIN-compliant flange)')
            ],
            'type': [
                ('2', 'Two-valve'),
                ('3', 'Three-valve'),
                ('5', 'Five-valve'),
                ('6', 'Five-valve natural gas metering pattern'),
                ('7', 'Two-valve (ASME B31.1 Power Piping)'),
                ('8', 'Three-valve (ASME B31.1 Power Piping)'),
                ('9', 'Five-valve (ASME B31.1 Power Piping)')
            ],
            'material': [
                ('2', '316 SST / 316L SST Body, 316 SST Stem/Tip'),
                ('3', 'Alloy C-276 Body, Bonnet, Stem/Tip'),
                ('4', 'Alloy 400 Body, Bonnet, Stem/Tip'),
                ('8', 'Alloy 625 Body, Bonnet, Stem/Tip'),
                ('9', 'Super Duplex SST (UNS S32760)')
            ],
            'conn': [
                ('A', '1/4-18 NPT female'),
                ('B', '1/2-14 NPT female'),
                ('S', '1/2-14 NPT female side entry (Coplanar only)')
            ],
            'packing': [
                ('1', 'PTFE'),
                ('2', 'Graphite-based')
            ],
            'seat': [
                ('1', 'Integral'),
                ('5', 'Soft POM (Only with NG pattern)')
            ]
        }

        self.options_data = {
            'Warranty': [('WR3', '3-year limited warranty'), ('WR5', '5-year limited warranty')],
            'Brackets': [
                ('B1', 'Bracket for 2-in. pipe, CS bolts'),
                ('B3', 'Flat bracket for 2-in. pipe, CS bolts'),
                ('B4', 'SST bracket for 2-in. pipe, 300 SST bolts'),
                ('B7', 'B1 bracket with 316 SST bolts'),
                ('B9', 'B3 bracket with 316 SST bolts'),
                ('BA', '316 SST B1 bracket with 316 SST bolts'),
                ('BC', '316 SST B3 bracket with 316 SST bolts'),
                ('BE', '316 SST B4 bracket with 316 SST bolts'),
                ('BF', 'CS panel mount bracket'),
                ('BG', '316 SST panel mount bracket')
            ],
            'Bolts': [
                ('L4', 'Austenitic 316 SST bolts'),
                ('L5', 'ASTM A193, Grade B7M bolts'),
                ('L8', 'ASTM A193, Grade B8M bolts, Class 2')
            ],
            'Cleaning': [('P2', 'Cleaning for special services')],
            'NACE': [
                ('SG', 'Sour gas (NACE MR0175/ISO 15156, MR0103)'),
                ('TI', 'Sour Gas materials with 316 SST non-wetted')
            ],
            'Certifications': [
                ('Q8', 'Material traceability cert per EN 10204 3.1'),
                ('Q15', 'NACE MR0175/ISO 15156 Compliance'),
                ('Q25', 'NACE MR0103 Compliance')
            ],
            'Adapters': [
                ('DF', '1/2-14 NPT female flange adapter'),
                ('DQ', '0.47-in. (12 mm) ferrule tube flange adapter')
            ],
            'Cold Temp': [
                ('CW1', '-67 °F (-55 °C) cold temp operation'),
                ('BR6', '-76 °F (-60 °C) cold temp operation')
            ],
            'Other': [
                ('PF', 'Relocated equalize valve for 9295'),
                ('HK', 'M10 process flange bolting'),
                ('HL', 'M12 process flange bolting'),
                ('DM', 'Drain Vent Plugs, 316 SST'),
                ('DS', 'Drain Vent Screen')
            ]
        }
        
        # Field Labels
        self.labels = {
            'style': 'Manifold Style',
            'type': 'Manifold Type',
            'material': 'Materials of Construction',
            'conn': 'Process Connection Style',
            'packing': 'Packing Material',
            'seat': 'Valve Seat'
        }

    def create_widgets(self):
        # Base Model Info
        base_frame = ttk.LabelFrame(self.scrollable_frame, text="Required Selections", padding="10")
        base_frame.pack(fill="x", pady=5)

        # Static Row for Model and Manufacturer
        static_frame = ttk.Frame(base_frame)
        static_frame.pack(fill="x", pady=2)
        ttk.Label(static_frame, text="Model: 0305", width=15, font=('Arial', 10, 'bold')).pack(side="left", padx=5)
        ttk.Label(static_frame, text="Manufacturer: R", width=15, font=('Arial', 10, 'bold')).pack(side="left", padx=5)

        # Dynamic Dropdowns
        self.combos = {}
        for key in self.data_keys:
            frame = ttk.Frame(base_frame)
            frame.pack(fill="x", pady=5)
            
            lbl = ttk.Label(frame, text=self.labels[key], width=25, anchor="w")
            lbl.pack(side="left")
            
            cb = ttk.Combobox(frame, textvariable=self.vars[key], state="readonly", width=60)
            cb.pack(side="left", fill="x", expand=True)
            cb.bind("<<ComboboxSelected>>", self.update_logic)
            self.combos[key] = cb
            
            # Populate initial values
            values = [f"{code} - {desc}" for code, desc in self.data_structure[key]]
            cb['values'] = values
            if values:
                cb.current(0)

        # Options Section
        opts_frame = ttk.LabelFrame(self.scrollable_frame, text="Additional Options", padding="10")
        opts_frame.pack(fill="x", pady=10)

        self.checkbuttons = {}
        
        # Grid layout for options
        row_idx = 0
        for category, items in self.options_data.items():
            ttk.Label(opts_frame, text=category, font=('Arial', 10, 'bold', 'underline')).grid(row=row_idx, column=0, sticky="w", pady=(10, 2))
            row_idx += 1
            
            col_idx = 0
            for code, desc in items:
                cb = ttk.Checkbutton(
                    opts_frame, 
                    text=f"{code} - {desc}", 
                    variable=self.option_vars[code],
                    command=self.update_result
                )
                cb.grid(row=row_idx, column=col_idx, sticky="w", padx=10, pady=2)
                self.checkbuttons[code] = cb
                
                col_idx += 1
                if col_idx > 1: # 2 columns of options
                    col_idx = 0
                    row_idx += 1
            if col_idx != 0:
                row_idx += 1

    def get_code(self, key):
        val = self.vars[key].get()
        return val.split(' - ')[0] if val else ''

    def update_logic(self, event=None):
        """
        Enforces constraints based on the selection guide notes.
        """
        current_style = self.get_code('style')
        current_type = self.get_code('type')
        current_mat = self.get_code('material')
        current_conn = self.get_code('conn')
        current_seat = self.get_code('seat')

        # --- Style Constraints ---
        # Type 5 not available with Traditional (T)
        # Type 6, 7, 8, 9 only available with Coplanar (C)
        valid_types = list(self.data_structure['type'])
        if current_style == 'T':
            valid_types = [t for t in valid_types if t[0] not in ['5', '6', '7', '8', '9']]
        elif current_style == 'M':
             # Note: PDF doesn't explicitly ban 6/7/8/9 for M, but usually NG/ASME are C or specific. 
             # Note 2 on Type says 6,7,8,9 Only available with coplanar C.
             valid_types = [t for t in valid_types if t[0] not in ['6', '7', '8', '9']]
        
        self.update_combo('type', valid_types)

        # --- Process Connection Constraints ---
        # A (1/4 NPT): Only available with Traditional (T) and M
        # B (1/2 NPT): Not available with M.
        # S (Side Entry): Only available with Coplanar (C), Types 2,3,5, Mat 2,3,4, Seat 1, Brackets B4/BE/SG
        valid_conns = list(self.data_structure['conn'])
        
        if current_style == 'C':
            # Remove A
            valid_conns = [c for c in valid_conns if c[0] != 'A']
            
            # Side Entry Logic (Complex)
            # Only avail with 2, 3, 5 type
            if current_type not in ['2', '3', '5']:
                valid_conns = [c for c in valid_conns if c[0] != 'S']
            # Only avail with 316, C-276, 400 (Codes 2, 3, 4)
            if current_mat not in ['2', '3', '4']:
                valid_conns = [c for c in valid_conns if c[0] != 'S']
            # Only integral seat
            if current_seat != '1':
                 valid_conns = [c for c in valid_conns if c[0] != 'S']

        elif current_style == 'T':
            # Remove S
            valid_conns = [c for c in valid_conns if c[0] != 'S']
            
        elif current_style == 'M':
            # Remove B and S
            valid_conns = [c for c in valid_conns if c[0] not in ['B', 'S']]

        self.update_combo('conn', valid_conns)

        # --- Material Constraints ---
        # Types 7, 8, 9, 6 only available with 316 SST (Code 2)
        if current_type in ['6', '7', '8', '9']:
            valid_mats = [m for m in self.data_structure['material'] if m[0] == '2']
        else:
            # Alloy 625 (8) and Duplex (9) only available with 2, 3, 5 valve type
            valid_mats = list(self.data_structure['material'])
            if current_type not in ['2', '3', '5']:
                valid_mats = [m for m in valid_mats if m[0] not in ['8', '9']]
        
        self.update_combo('material', valid_mats)

        # --- Seat Constraints ---
        # Soft POM (5) only available with Natural Gas (Type 6)
        valid_seats = list(self.data_structure['seat'])
        if current_type != '6':
            valid_seats = [s for s in valid_seats if s[0] != '5']
        self.update_combo('seat', valid_seats)

        # --- Option Constraints (Visual disable) ---
        # Adapters (DF, DQ): Only with T and M
        state = 'normal' if current_style in ['T', 'M'] else 'disabled'
        self.toggle_option('DF', state)
        self.toggle_option('DQ', state)
        if state == 'disabled':
            self.option_vars['DF'].set(False)
            self.option_vars['DQ'].set(False)

        self.update_result()

    def update_combo(self, key, valid_values):
        """
        Updates a combobox values if they differ from current.
        Preserves selection if valid, else resets to first valid.
        """
        cb = self.combos[key]
        current_val = self.vars[key].get()
        current_code = current_val.split(' - ')[0] if current_val else ''
        
        formatted_values = [f"{code} - {desc}" for code, desc in valid_values]
        
        if cb['values'] != tuple(formatted_values):
            cb['values'] = formatted_values
            
            # Check if current selection is still valid
            found = False
            for val in formatted_values:
                if val.startswith(current_code + ' -'):
                    cb.set(val)
                    found = True
                    break
            
            if not found and formatted_values:
                cb.current(0)
            elif not formatted_values:
                cb.set('')

    def toggle_option(self, code, state):
        if code in self.checkbuttons:
            self.checkbuttons[code].configure(state=state)

    def update_result(self):
        # Build the model string
        components = [
            "0305",
            "R",
            self.get_code('style'),
            self.get_code('type'),
            self.get_code('material'),
            self.get_code('conn'),
            self.get_code('packing'),
            self.get_code('seat')
        ]
        
        # Add selected options
        opts = []
        for code, var in self.option_vars.items():
            if var.get():
                opts.append(code)
        
        # Combine
        full_model = " ".join(components)
        if opts:
            full_model += " " + " ".join(opts)
            
        self.result_label.config(text=full_model)

    def copy_to_clipboard(self):
        self.clipboard_clear()
        self.clipboard_append(self.result_label.cget("text"))
        messagebox.showinfo("Copied", "Model number copied to clipboard!")

if __name__ == "__main__":
    app = Rosemount305Configurator()
    app.mainloop()