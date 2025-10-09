import tkinter as tk
from tkinter import ttk, scrolledtext
import re

class GaugeConfigurator:
    def __init__(self, root):
        self.root = root
        self.root.title("Ashcroft 1279 Duragauge Model Number Configurator")
        self.root.geometry("900x700")
        
        # Configuration data
        self.setup_options()
        
        # Create main container with scrollbar
        main_frame = ttk.Frame(root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Title
        title = ttk.Label(main_frame, text="1279 Duragauge Pressure Gauge Configurator", 
                         font=('Arial', 14, 'bold'))
        title.grid(row=0, column=0, columnspan=2, pady=10)
        
        # Create canvas and scrollbar for scrollable content
        canvas = tk.Canvas(main_frame, highlightthickness=0)
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        scrollbar.grid(row=1, column=2, sticky=(tk.N, tk.S))
        
        main_frame.rowconfigure(1, weight=1)
        
        # Store dropdown variables
        self.vars = {}
        
        # Create dropdowns
        row = 0
        self.create_dropdown(scrollable_frame, row, "Dial Size/Model", "dial_size", self.dial_sizes)
        row += 1
        self.create_dropdown(scrollable_frame, row, "System (Tube/Connection)", "system", self.systems)
        row += 1
        self.create_dropdown(scrollable_frame, row, "Case Design", "case_design", self.case_designs)
        row += 1
        self.create_dropdown(scrollable_frame, row, "Process Connection Size", "connection_size", self.connection_sizes)
        row += 1
        self.create_dropdown(scrollable_frame, row, "Process Connection Location", "connection_location", self.connection_locations)
        row += 1
        
        # Options section
        ttk.Label(scrollable_frame, text="Options (Select all that apply):", 
                 font=('Arial', 10, 'bold')).grid(row=row, column=0, columnspan=2, pady=(15,5), sticky=tk.W)
        row += 1
        
        self.option_vars = {}
        for opt_code, opt_desc in self.options.items():
            var = tk.BooleanVar()
            self.option_vars[opt_code] = var
            cb = ttk.Checkbutton(scrollable_frame, text=f"{opt_code} - {opt_desc}", variable=var,
                                command=self.update_model_number)
            cb.grid(row=row, column=0, columnspan=2, sticky=tk.W, padx=20)
            row += 1
        
        # Range selection
        ttk.Label(scrollable_frame, text="", font=('Arial', 1)).grid(row=row, column=0)
        row += 1
        self.create_dropdown(scrollable_frame, row, "Pressure Range", "range", self.ranges)
        
        # Separator
        ttk.Separator(main_frame, orient='horizontal').grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=10)
        
        # Model number display
        result_frame = ttk.LabelFrame(main_frame, text="Generated Model Number", padding="10")
        result_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        result_frame.columnconfigure(0, weight=1)
        
        self.model_number_var = tk.StringVar(value="451279 S SH 04 L XLL 15#")
        model_entry = ttk.Entry(result_frame, textvariable=self.model_number_var, 
                               font=('Courier', 12, 'bold'), state='readonly')
        model_entry.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=5)
        
        copy_btn = ttk.Button(result_frame, text="Copy to Clipboard", command=self.copy_to_clipboard)
        copy_btn.grid(row=0, column=1, padx=5)
        
        # Summary text
        summary_frame = ttk.LabelFrame(main_frame, text="Configuration Summary", padding="10")
        summary_frame.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        summary_frame.columnconfigure(0, weight=1)
        summary_frame.rowconfigure(0, weight=1)
        
        self.summary_text = scrolledtext.ScrolledText(summary_frame, height=8, width=80, 
                                                      wrap=tk.WORD, font=('Arial', 9))
        self.summary_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        main_frame.rowconfigure(4, weight=1)
        
        # Bind mousewheel to canvas
        canvas.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(int(-1*(e.delta/120)), "units"))
        
        # Initial update
        self.update_model_number()
    
    def setup_options(self):
        self.dial_sizes = [("451279", "4½″ solid front")]
        
        self.systems = [
            ("A", "Bronze tube, brass process connection (Max 1,000 psi)"),
            ("P", "K-Monel 500 tube, Monel 400 process connection (Max 30,000 psi)"),
            ("R", "316L SS tube, steel process connection (Max 20,000 psi)"),
            ("S", "316 SS tube, 316L SS process connection (Max 20,000 psi, NSF-61)")
        ]
        
        self.case_designs = [
            ("S", "Solid front case, dry"),
            ("SH", "Solid front case, dry (IP66 NEMA 4X)"),
            ("SL", "Solid front case, liquid filled (glycerin STD.)")
        ]
        
        self.connection_sizes = [
            ("02", "¼ NPT Male (N/A for ranges over 20,000 psi)"),
            ("04", "½ NPT Male (N/A for ranges over 20,000 psi)"),
            ("09", "⁹⁄₁₆ 18 UNF-2B, Aminco high pressure (over 20,000 psi)"),
            ("AM", "AND 10050-4 (¼ tubing connection)"),
            ("RW", "SAE ⁷⁄₁₆ & 20 Straight thread")
        ]
        
        self.connection_locations = [
            ("L", "Lower"),
            ("B", "Back"),
            ("D", "Side (3 o'clock)"),
            ("E", "Side connection (9 o'clock)"),
            ("T", "Top connection")
        ]
        
        self.options = {
            "GV": "Silicone case fill",
            "GX": "Halocarbon case fill",
            "LL": "PLUS! Performance",
            "ND": "PLUS! Performance, silicone free (glycerin/silicone filled)",
            "NZ": "PLUS! Performance, silicone free (dry gauges)",
            "X1P": "IP65 dry case (cannot be liquid filled)",
            "NH": "SS tag wired to case",
            "PD": "Acrylic window",
            "SG": "Safety Glass",
            "C4": "Traceable calibration certificate",
            "6B": "Cleaned for oxygen service",
            "5G": "Attach one accessory to gauge",
            "AB": "Calibrated for absolute pressure",
            "C3": "Material Traceability Report (EN 10204 3.1)",
            "DA": "Dial marking (text on dial)",
            "D3": "DuraVis retroreflective dial",
            "EP": "Maximum pointer (N/A with liquid/hermetic)",
            "HY": "Hydrostatic/pneumatic testing",
            "MQ": "Positive Material Identification",
            "NG": "Non-glare glass",
            "OS": "Overload stop",
            "SH": "Red set hand, stationary",
            "TM": "2″ Pipe mounting bracket",
            "TS": "Throttle screw",
            "VS": "Underload stop",
            "VY": "Krytox Lubricated Movement (-65°F/-54°C)",
            "56": "Flush mounting ring"
        }
        
        self.ranges = [
            ("30IMV", "30″ Hg Vacuum"),
            ("V/15#", "Vacuum/15 psi compound"),
            ("V/30#", "Vacuum/30 psi compound"),
            ("V/60#", "Vacuum/60 psi compound"),
            ("V/100#", "Vacuum/100 psi compound"),
            ("15#", "15 psi"),
            ("20#", "20 psi"),
            ("30#", "30 psi"),
            ("60#", "60 psi"),
            ("100#", "100 psi"),
            ("120#", "120 psi"),
            ("160#", "160 psi"),
            ("200#", "200 psi"),
            ("300#", "300 psi"),
            ("400#", "400 psi"),
            ("500#", "500 psi"),
            ("600#", "600 psi"),
            ("800#", "800 psi"),
            ("1000#", "1000 psi"),
            ("1500#", "1500 psi"),
            ("2000#", "2000 psi"),
            ("3000#", "3000 psi"),
            ("4000#", "4000 psi"),
            ("5000#", "5000 psi"),
            ("6000#", "6000 psi"),
            ("8000#", "8000 psi"),
            ("10000#", "10000 psi"),
            ("15000#", "15000 psi"),
            ("20000#", "20000 psi"),
            ("30000#", "30000 psi"),
            ("1BR", "1 bar"),
            ("1.6BR", "1.6 bar"),
            ("2.5BR", "2.5 bar"),
            ("4BR", "4 bar"),
            ("6BR", "6 bar"),
            ("10BR", "10 bar"),
            ("16BR", "16 bar"),
            ("25BR", "25 bar"),
            ("40BR", "40 bar"),
            ("60BR", "60 bar"),
            ("100BR", "100 bar"),
            ("160BR", "160 bar"),
            ("250BR", "250 bar"),
            ("400BR", "400 bar"),
            ("600BR", "600 bar"),
            ("1000BR", "1000 bar"),
            ("1600BR", "1600 bar")
        ]
    
    def create_dropdown(self, parent, row, label_text, var_name, options):
        label = ttk.Label(parent, text=label_text + ":", font=('Arial', 9, 'bold'))
        label.grid(row=row, column=0, sticky=tk.W, pady=5, padx=5)
        
        var = tk.StringVar()
        self.vars[var_name] = var
        
        # Create display values
        display_options = [f"{code} - {desc}" for code, desc in options]
        
        combo = ttk.Combobox(parent, textvariable=var, values=display_options, 
                            state='readonly', width=60)
        combo.grid(row=row, column=1, sticky=(tk.W, tk.E), pady=5, padx=5)
        combo.current(0)
        combo.bind('<<ComboboxSelected>>', lambda e: self.update_model_number())
        
        parent.columnconfigure(1, weight=1)
    
    def extract_code(self, full_value):
        """Extract the code portion from 'CODE - Description' format"""
        if ' - ' in full_value:
            return full_value.split(' - ')[0]
        return full_value
    
    def update_model_number(self):
        # Extract codes from selections
        parts = []
        parts.append(self.extract_code(self.vars['dial_size'].get()))
        parts.append(self.extract_code(self.vars['system'].get()))
        parts.append(self.extract_code(self.vars['case_design'].get()))
        parts.append(self.extract_code(self.vars['connection_size'].get()))
        parts.append(self.extract_code(self.vars['connection_location'].get()))
        
        # Add options
        selected_options = [code for code, var in self.option_vars.items() if var.get()]
        if selected_options:
            parts.append("X" + "".join(selected_options))
        else:
            parts.append("")  # No options selected
        
        parts.append(self.extract_code(self.vars['range'].get()))
        
        # Create model number
        model_number = " ".join(parts)
        self.model_number_var.set(model_number)
        
        # Update summary
        self.update_summary()
    
    def update_summary(self):
        self.summary_text.delete(1.0, tk.END)
        
        summary = "Configuration Details:\n" + "="*70 + "\n\n"
        
        for var_name, var in self.vars.items():
            label = var_name.replace('_', ' ').title()
            value = var.get()
            summary += f"{label}: {value}\n"
        
        selected_options = [f"{code} - {self.options[code]}" 
                          for code, var in self.option_vars.items() if var.get()]
        if selected_options:
            summary += f"\nSelected Options:\n"
            for opt in selected_options:
                summary += f"  • {opt}\n"
        else:
            summary += f"\nSelected Options: None\n"
        
        self.summary_text.insert(1.0, summary)
    
    def copy_to_clipboard(self):
        self.root.clipboard_clear()
        self.root.clipboard_append(self.model_number_var.get())
        
        # Show feedback
        original_text = self.model_number_var.get()
        self.model_number_var.set("✓ Copied to clipboard!")
        self.root.after(1500, lambda: self.model_number_var.set(original_text))

if __name__ == "__main__":
    root = tk.Tk()
    app = GaugeConfigurator(root)
    root.mainloop()