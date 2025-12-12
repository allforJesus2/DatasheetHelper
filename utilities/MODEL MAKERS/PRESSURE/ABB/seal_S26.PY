import tkinter as tk
from tkinter import ttk, messagebox
import re

class SealConfigurator:
    def __init__(self, root):
        self.root = root
        self.root.title("S26 Diaphragm Seal Model Number Configurator")
        self.root.geometry("900x700")
        
        # Model configurations
        self.models = {
            "S26RA": "Rotating flange (flush/extended) - ASME B16.5",
            "S26RE": "Rotating flange (flush/extended) - EN 1092-1",
            "S26RJ": "Rotating flange (flush) - JIS",
            "S26RR": "Rotating flange (flush) Ring Joint - ASME B16.5",
            "S26RH": "Rotating flange (flush) - ISO 10423",
            "S26FA": "Fixed flange (flush/extended) - ASME B16.5",
            "S26FE": "Fixed flange (flush/extended) - EN 1092-1",
            "S26MA": "Off-line flange - ASME B16.5",
            "S26ME": "Off-line flange - EN 1092-1",
            "S26TT": "Off-line threaded",
            "S26SS": "Sanitary and food",
            "S26WA": "Wafer - ASME B16.5",
            "S26WE": "Wafer - EN 1092-1",
            "S26BN": "Button type remote",
            "S26PN": "Urea service remote",
            "S26KN": "Pulp and paper"
        }
        
        self.variables = {}
        self.dropdowns = {}
        
        self.create_widgets()
        
    def create_widgets(self):
        # Main container with canvas and scrollbar
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Canvas
        canvas = tk.Canvas(main_frame)
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Title
        title = ttk.Label(scrollable_frame, text="S26 Diaphragm Seal Configurator", 
                         font=('Arial', 16, 'bold'))
        title.grid(row=0, column=0, columnspan=3, pady=10)
        
        # Base Model Selection
        row = 1
        ttk.Label(scrollable_frame, text="Base Model:", font=('Arial', 10, 'bold')).grid(
            row=row, column=0, sticky='w', padx=5, pady=5)
        self.variables['base'] = tk.StringVar()
        base_dropdown = ttk.Combobox(scrollable_frame, textvariable=self.variables['base'], 
                                     width=50, state='readonly')
        base_dropdown['values'] = [f"{k} - {v}" for k, v in self.models.items()]
        base_dropdown.grid(row=row, column=1, columnspan=2, sticky='ew', padx=5, pady=5)
        base_dropdown.bind('<<ComboboxSelected>>', self.on_model_change)
        self.dropdowns['base'] = base_dropdown
        
        # Configuration Frame
        self.config_frame = ttk.LabelFrame(scrollable_frame, text="Configuration Options", 
                                          padding=10)
        self.config_frame.grid(row=2, column=0, columnspan=3, sticky='ew', padx=5, pady=10)
        
        # Result Frame
        result_frame = ttk.LabelFrame(scrollable_frame, text="Generated Model Number", 
                                     padding=10)
        result_frame.grid(row=3, column=0, columnspan=3, sticky='ew', padx=5, pady=10)
        
        self.result_var = tk.StringVar(value="Select base model to begin")
        result_label = ttk.Label(result_frame, textvariable=self.result_var, 
                                font=('Courier', 14, 'bold'), foreground='blue')
        result_label.pack()
        
        # Buttons
        button_frame = ttk.Frame(scrollable_frame)
        button_frame.grid(row=4, column=0, columnspan=3, pady=10)
        
        ttk.Button(button_frame, text="Generate Model Number", 
                  command=self.generate_model).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Copy to Clipboard", 
                  command=self.copy_to_clipboard).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Reset", 
                  command=self.reset_form).pack(side=tk.LEFT, padx=5)
        
        # Pack canvas and scrollbar
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Mouse wheel scrolling
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        
    def on_model_change(self, event=None):
        # Clear previous configuration
        for widget in self.config_frame.winfo_children():
            widget.destroy()
        
        selected = self.variables['base'].get().split(' - ')[0]
        
        if selected == "S26RA":
            self.create_s26ra_config()
        elif selected == "S26RE":
            self.create_s26re_config()
        elif selected == "S26RJ":
            self.create_s26rj_config()
        elif selected == "S26FA":
            self.create_s26fa_config()
        elif selected == "S26SS":
            self.create_s26ss_config()
        elif selected == "S26WA":
            self.create_s26wa_config()
        elif selected == "S26MA":
            self.create_s26ma_config()
        else:
            ttk.Label(self.config_frame, text="Configuration options for this model coming soon...",
                     font=('Arial', 10, 'italic')).pack()
            
    def create_dropdown(self, parent, row, label, var_name, options):
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky='w', padx=5, pady=3)
        self.variables[var_name] = tk.StringVar()
        dropdown = ttk.Combobox(parent, textvariable=self.variables[var_name], 
                               width=60, state='readonly')
        dropdown['values'] = options
        dropdown.grid(row=row, column=1, sticky='ew', padx=5, pady=3)
        dropdown.current(0)
        self.dropdowns[var_name] = dropdown
        return dropdown
        
    def create_s26ra_config(self):
        options_config = [
            ("Transmitter Side:", "transmitter", ["H - High pressure side", "L - Low pressure side"]),
            ("Mounting Flange:", "flange", [
                "E1 - ASME CL 150 / 2 in.", "E2 - ASME CL 300 / 2 in.", 
                "E3 - ASME CL 600 / 2 in.", "G1 - ASME CL 150 / 3 in.",
                "H1 - ASME CL 150 / 4 in."
            ]),
            ("Flange Material:", "flange_mat", [
                "C - Carbon steel", "S - Dual Grade AISI 316 ss"
            ]),
            ("Extension:", "extension", [
                "F - Flush", "1 - 50mm AISI 316L ss", "3 - 100mm AISI 316L ss",
                "5 - 150mm AISI 316L ss"
            ]),
            ("Diaphragm Material:", "diaphragm", [
                "SM - AISI 316L ss (NACE)", "SL - AISI 316L ss Low thickness (NACE)",
                "HM - Hastelloy C-276 (NACE)", "HL - Hastelloy C-276 Low thickness (NACE)",
                "MM - Hastelloy C-2000 (NACE)", "KM - AISI 316L ss PFA anti-stick (NACE)"
            ]),
            ("Seal Surface:", "surface", [
                "1 - Serrated", "2 - Smooth"
            ]),
            ("Capillary Protection:", "cap_protect", [
                "A - AISI 316L ss armour", "B - AISI 316L ss armour with PVC cover",
                "N - Extension tube for direct mount"
            ]),
            ("Capillary Length:", "cap_length", [
                "1 - Direct-mount", "A - 1m (3ft)", "B - 1.5m (5ft)", "C - 2m (7ft)",
                "E - 3m (10ft)", "J - 5m (17ft)", "L - 6m (20ft)", "S - 10m (33ft)"
            ]),
            ("Fill Fluid:", "fill", [
                "S - Silicone oil PMX 200 10 cSt", "P - Silicone oil Baysilone PD5 5 cSt",
                "N - Inert oil Galden G5 (Oxygen)", "G - Silicone oil high temp",
                "W - White oil (FDA)", "A - Vegetable oil Neobee M-20 (FDA)"
            ]),
            ("Flushing Ring Holes:", "flush_holes", [
                "N - None", "2 - 1 hole 1/2 in. NPT", "3 - 2 holes 1/2 in. NPT",
                "4 - 1 hole 1/4 in. NPT", "5 - 2 holes 1/4 in. NPT"
            ]),
            ("Flushing Ring Material:", "flush_mat", [
                "N - None", "A - AISI 316 ss (NACE)", "H - Hastelloy C-276 (NACE)"
            ]),
            ("Flushing Ring Plug:", "flush_plug", [
                "N - No plug - No gasket", "A - No plug - garlock", "B - No plug - PTFE",
                "D - AISI 316L ss - no gasket (NACE)", "E - AISI 316L ss - garlock (NACE)",
                "F - AISI 316L ss - PTFE (NACE)"
            ])
        ]
        
        for i, (label, var_name, options) in enumerate(options_config):
            self.create_dropdown(self.config_frame, i, label, var_name, options)
            
    def create_s26re_config(self):
        options_config = [
            ("Transmitter Side:", "transmitter", ["H - High pressure side", "L - Low pressure side"]),
            ("Mounting Flange:", "flange", [
                "N2 - PN 16-40 / DN 50", "N3 - PN 63 / DN 50", "P1 - PN 16 / DN 80",
                "Q1 - PN 16 / DN 100"
            ]),
            ("Flange Material:", "flange_mat", ["C - Carbon steel", "S - AISI 316 ss"]),
            ("Extension:", "extension", [
                "F - Flush", "1 - 50mm AISI 316L ss", "3 - 100mm AISI 316L ss"
            ]),
            ("Diaphragm Material:", "diaphragm", [
                "SM - AISI 316L ss (NACE)", "HM - Hastelloy C-276 (NACE)",
                "MM - Hastelloy C-2000 (NACE)", "KM - AISI 316L ss PFA anti-stick (NACE)"
            ]),
            ("Seal Surface:", "surface", ["1 - Serrated", "2 - Smooth"]),
            ("Capillary Protection:", "cap_protect", [
                "A - AISI 316L ss armour", "B - AISI 316L ss armour with PVC",
                "N - Extension tube for direct mount"
            ]),
            ("Capillary Length:", "cap_length", [
                "1 - Direct-mount", "A - 1m (3ft)", "C - 2m (7ft)", "E - 3m (10ft)",
                "L - 6m (20ft)"
            ]),
            ("Fill Fluid:", "fill", [
                "S - Silicone oil PMX 200", "P - Silicone oil Baysilone PD5",
                "N - Inert oil Galden G5", "W - White oil (FDA)"
            ]),
            ("Flushing Ring Holes:", "flush_holes", [
                "N - None", "2 - 1 hole 1/2 in. NPT", "3 - 2 holes 1/2 in. NPT"
            ]),
            ("Flushing Ring Material:", "flush_mat", [
                "N - None", "A - AISI 316L ss (NACE)", "H - Hastelloy C-276 (NACE)"
            ]),
            ("Flushing Ring Plug:", "flush_plug", [
                "N - No plug", "D - AISI 316L ss no gasket (NACE)",
                "E - AISI 316L ss garlock (NACE)"
            ])
        ]
        
        for i, (label, var_name, options) in enumerate(options_config):
            self.create_dropdown(self.config_frame, i, label, var_name, options)
            
    def create_s26rj_config(self):
        options_config = [
            ("Transmitter Side:", "transmitter", ["H - High pressure side", "L - Low pressure side"]),
            ("Mounting Flange:", "flange", [
                "B2 - 10K / A50", "B4 - 20K / A50", "C2 - 10K / A80", "D2 - 10K / A100"
            ]),
            ("Flange Material:", "flange_mat", ["C - Carbon steel", "S - AISI 316 ss"]),
            ("Extension:", "extension", ["F - Flush"]),
            ("Diaphragm Material:", "diaphragm", [
                "SM - AISI 316L ss (NACE)", "HM - Hastelloy C-276 (NACE)",
                "MM - Hastelloy C-2000 (NACE)"
            ]),
            ("Seal Surface:", "surface", ["1 - Serrated", "2 - Smooth"]),
            ("Capillary Protection:", "cap_protect", [
                "A - AISI 316L ss armour", "B - AISI 316L ss armour with PVC",
                "N - Extension tube"
            ]),
            ("Capillary Length:", "cap_length", [
                "1 - Direct-mount", "A - 1m (3ft)", "E - 3m (10ft)", "L - 6m (20ft)"
            ]),
            ("Fill Fluid:", "fill", [
                "S - Silicone oil PMX 200", "P - Silicone oil Baysilone PD5",
                "N - Inert oil Galden G5"
            ]),
            ("Flushing Ring:", "flush_holes", ["N - None"]),
            ("Flushing Material:", "flush_mat", ["N - None"]),
            ("Flushing Plug:", "flush_plug", ["N - None"])
        ]
        
        for i, (label, var_name, options) in enumerate(options_config):
            self.create_dropdown(self.config_frame, i, label, var_name, options)
            
    def create_s26fa_config(self):
        options_config = [
            ("Transmitter Side:", "transmitter", ["H - High pressure side", "L - Low pressure side"]),
            ("Mounting Flange:", "flange", [
                "E1 - ASME CL 150 / 2 in.", "E2 - ASME CL 300 / 2 in.",
                "G1 - ASME CL 150 / 3 in.", "H1 - ASME CL 150 / 4 in."
            ]),
            ("Flange Material:", "flange_mat", ["S - Dual Grade AISI 316 ss"]),
            ("Extension:", "extension", [
                "F - Flush", "1 - 50mm AISI 316L ss", "3 - 100mm AISI 316L ss"
            ]),
            ("Diaphragm Material:", "diaphragm", [
                "SM - AISI 316L ss (NACE)", "SL - AISI 316L ss Low thickness (NACE)",
                "HM - Hastelloy C-276 (NACE)", "MM - Hastelloy C-2000 (NACE)"
            ]),
            ("Seal Surface:", "surface", ["1 - Serrated", "2 - Smooth"]),
            ("Capillary Protection:", "cap_protect", [
                "A - AISI 316L ss armour", "B - AISI 316L ss armour with PVC",
                "N - Extension tube"
            ]),
            ("Capillary Length:", "cap_length", [
                "1 - Direct-mount", "A - 1m (3ft)", "E - 3m (10ft)", "S - 10m (33ft)"
            ]),
            ("Fill Fluid:", "fill", [
                "S - Silicone oil PMX 200", "P - Silicone oil Baysilone PD5",
                "W - White oil (FDA)"
            ]),
            ("Flushing Ring Holes:", "flush_holes", ["N - None", "2 - 1 hole 1/2 in. NPT"]),
            ("Flushing Ring Material:", "flush_mat", ["N - None", "A - AISI 316 ss (NACE)"]),
            ("Flushing Ring Plug:", "flush_plug", ["N - No plug", "D - AISI 316L ss (NACE)"])
        ]
        
        for i, (label, var_name, options) in enumerate(options_config):
            self.create_dropdown(self.config_frame, i, label, var_name, options)
            
    def create_s26ss_config(self):
        options_config = [
            ("Transmitter Side:", "transmitter", ["H - High pressure side", "L - Low pressure side"]),
            ("Mounting Connection:", "mounting", [
                "F - 2 in Triclamp", "G - 3 in Triclamp", "H - 4 in Triclamp",
                "L - 2 in Cherry Burrell", "P - 4 in Sanitary flush"
            ]),
            ("Diaphragm Material:", "diaphragm", ["SM - AISI 316L ss"]),
            ("Capillary Protection:", "cap_protect", [
                "A - AISI 316L ss armour", "B - AISI 316L ss armour with PVC",
                "N - Extension tube"
            ]),
            ("Capillary Length:", "cap_length", [
                "1 - Direct-mount", "A - 1m (3ft)", "E - 3m (10ft)", "L - 6m (20ft)"
            ]),
            ("Fill Fluid:", "fill", [
                "S - Silicone oil PMX 200", "C - Silicone polymer Syltherm XLT",
                "W - White oil (FDA)", "A - Vegetable oil Neobee M-20 (FDA)"
            ]),
            ("Clamp/Fittings:", "clamp", [
                "1 - None", "A - 2 in V-band Clamp", "B - 3 in V-band Clamp",
                "C - 4 in V-band Clamp"
            ]),
            ("Gasket:", "gasket", [
                "1 - None", "A - Ethylene propylene DN100 (EPDM)",
                "G - Ethylene propylene (EPDM)"
            ])
        ]
        
        for i, (label, var_name, options) in enumerate(options_config):
            self.create_dropdown(self.config_frame, i, label, var_name, options)
            
    def create_s26wa_config(self):
        options_config = [
            ("Transmitter Side:", "transmitter", ["H - High pressure side", "L - Low pressure side"]),
            ("Mounting Flange:", "flange", [
                "D5 - ASME 1 1/2 in.", "E5 - ASME 2 in.", "G5 - ASME 3 in."
            ]),
            ("Extension:", "extension", ["F - Flush"]),
            ("Diaphragm Material:", "diaphragm", [
                "SM - AISI 316L ss (NACE)", "SL - AISI 316L ss Low thickness (NACE)",
                "HM - Hastelloy C-276 (NACE)", "MM - Hastelloy C-2000 (NACE)"
            ]),
            ("Seal Surface:", "surface", ["1 - Serrated", "2 - Smooth"]),
            ("Capillary Protection:", "cap_protect", [
                "A - AISI 316L ss armour", "B - AISI 316L ss armour with PVC"
            ]),
            ("Capillary Length:", "cap_length", [
                "A - 1m (3ft)", "C - 2m (7ft)", "E - 3m (10ft)", "L - 6m (20ft)"
            ]),
            ("Fill Fluid:", "fill", [
                "S - Silicone oil PMX 200", "P - Silicone oil Baysilone PD5",
                "N - Inert oil Galden G5"
            ]),
            ("Flushing Ring Holes:", "flush_holes", ["N - None", "2 - 1 hole 1/2 in. NPT"]),
            ("Flushing Ring Material:", "flush_mat", ["N - None", "A - AISI 316L ss (NACE)"]),
            ("Flushing Ring Plug:", "flush_plug", ["N - No plug", "D - AISI 316L ss (NACE)"])
        ]
        
        for i, (label, var_name, options) in enumerate(options_config):
            self.create_dropdown(self.config_frame, i, label, var_name, options)
            
    def create_s26ma_config(self):
        options_config = [
            ("Transmitter Side:", "transmitter", ["H - High pressure side", "L - Low pressure side"]),
            ("Mounting Flange:", "flange", [
                "A1 - ASME CL 150 / 1/2 in.", "A2 - ASME CL 300 / 1/2 in.",
                "C1 - ASME CL 150 / 1 in.", "C2 - ASME CL 300 / 1 in.",
                "D1 - ASME CL 150 / 1 1/2 in.", "D2 - ASME CL 300 / 1 1/2 in."
            ]),
            ("Flange Material / Seat:", "flange_mat", [
                "S - AISI 316L ss / Form RF serrated (NACE)",
                "H - Hastelloy C-276 / Form RF serrated (NACE)",
                "Y - Hastelloy C-2000 / Form RF serrated (NACE)"
            ]),
            ("Diaphragm Material:", "diaphragm", [
                "SM - AISI 316L ss (NACE)", "HM - Hastelloy C-276 (NACE)",
                "MM - Hastelloy C-2000 (NACE)", "ZM - Hastelloy C-2000 diaphragm and body (NACE)",
                "LM - Inconel 625 (NACE)", "TM - Tantalum", "NM - AISI 316L ss gold plated (NACE)"
            ]),
            ("Capillary Protection:", "cap_protect", [
                "A - AISI 316L ss armour", "B - AISI 316L ss armour with PVC protective cover",
                "N - Extension tube for direct mount seal"
            ]),
            ("Capillary Length:", "cap_length", [
                "1 - Direct-mount construction", "A - 1m (3ft)", "B - 1.5m (5ft)", 
                "C - 2m (7ft)", "D - 2.5m (8ft)", "E - 3m (10ft)", "F - 3.5m (12ft)",
                "G - 4m (13ft)", "H - 4.5m (15ft)", "J - 5m (17ft)", "K - 5.5m (18ft)",
                "L - 6m (20ft)", "M - 6.5m (22ft)", "N - 7m (23.5ft)", "P - 7.5m (25ft)",
                "Q - 8m (27ft)", "R - 9m (30ft)", "S - 10m (33ft)", "T - 12m (40ft)"
            ]),
            ("Fill Fluid:", "fill", [
                "S - Silicone oil PMX 200 10 cSt (-40 to 250°C)", 
                "P - Silicone oil Baysilone PD5 5 cSt (-85 to 250°C)",
                "N - Inert oil Galden G5 (Oxygen service)",
                "G - Silicone oil for high temperature (-10 to 375°C)",
                "C - Silicone polymer Syltherm XLT (-100 to 100°C)",
                "W - White oil (FDA approved)",
                "A - Vegetable oil Neobee M-20 (FDA approved)",
                "B - Glycerin-water 70% (FDA approved)"
            ]),
            ("Flushing Connections:", "flushing", [
                "1 - Not required", "Q - Provided (with 2 plugs supplied)"
            ]),
            ("Gasket:", "gasket", [
                "2 - PTFE", "3 - Viton®", "7 - Graphite"
            ])
        ]
        
        for i, (label, var_name, options) in enumerate(options_config):
            self.create_dropdown(self.config_frame, i, label, var_name, options)
            
    def generate_model(self):
        if 'base' not in self.variables or not self.variables['base'].get():
            messagebox.showwarning("Warning", "Please select a base model first!")
            return
            
        base_model = self.variables['base'].get().split(' - ')[0]
        model_number = base_model
        
        # Extract codes from selections
        field_order = ['transmitter', 'flange', 'flange_mat', 'extension', 'diaphragm',
                      'surface', 'cap_protect', 'cap_length', 'fill', 'flush_holes',
                      'flush_mat', 'flush_plug', 'mounting', 'clamp', 'gasket', 'flushing']
        
        for field in field_order:
            if field in self.variables and self.variables[field].get():
                value = self.variables[field].get()
                # Extract code (everything before first space and dash)
                code = value.split(' - ')[0].split(' ')[0]
                model_number += code
                
        self.result_var.set(model_number)
        
    def copy_to_clipboard(self):
        model = self.result_var.get()
        if model and model != "Select base model to begin":
            self.root.clipboard_clear()
            self.root.clipboard_append(model)
            messagebox.showinfo("Success", f"Model number copied to clipboard:\n{model}")
        else:
            messagebox.showwarning("Warning", "No model number to copy!")
            
    def reset_form(self):
        for var in self.variables.values():
            var.set('')
        self.result_var.set("Select base model to begin")
        for widget in self.config_frame.winfo_children():
            widget.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = SealConfigurator(root)
    root.mainloop()
