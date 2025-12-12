import tkinter as tk
from tkinter import ttk, messagebox

class DurcoValveGenerator:
    def __init__(self, root):
        self.root = root
        self.root.title("Durco G4 Valve Model Number Generator")
        self.root.geometry("900x700")
        
        # Data structures based on the document
        self.valve_sizes = ["0.25", "0.38", "0.5", "0.75", "1", "1.5", "2", "3", "4", "5", "6", "8", "10", "12", "14", "16", "18", "20"]
        
        self.product_families = {
            "2 WAY G4": "G4",
            "3 WAY MG4": "MG4",
            "Marathon G4B": "G4B",
            "HF Alkylation Valve G4HF": "G4HF",
            "Full Jacket G4 FJG4": "FJG4",
            "Partial Jacket G4 PJG4": "PJG4",
            "2 Way* G4N": "G4N",
            "2 Way Acetic Acid G4AA": "G4AA"
        }
        
        self.pressure_classes = {
            "150": "1",
            "PN 10": "2",
            "300": "3",
            "PN 16": "4",
            "300x600": "5",
            "600": "6",
            "PN 25": "7",
            "PN 40": "8"
        }
        
        self.end_configs = {
            "Raised Face Flange": "1",
            "Screwed End": "2",
            "Butt Weld": "3",
            "Socket Weld": "4",
            "Buttweld X Socket Weld": "5",
            "Flat Face Flange": "F",
            "Ring Type Joint": "R",
            "Small Groove Flange": "S",
            "Large Groove Flange": "T"
        }
        
        self.plug_styles = {
            "2 Way": "C",
            "3 Way Arr 1": "D",
            "3 Way Arr 3": "E",
            "3 Way Arr 5": "F",
            "3 Way Arr 7": "G",
            "3 Way Arr 8": "H",
            "3 Way Arr 13": "J",
            "Cv 4 1\"": "K",
            "Cv 8 1\"": "L",
            "Cv 30 1\"": "M",
            "Cv 31 1.5\"": "N",
            "Cv 54 2\"": "P",
            "Cv 12 3\"": "Q",
            "Cv 188 4\"": "R",
            "Cv 370 6\"": "S",
            "Cv 3 1\"": "U",
            "Cv 1 1\"": "W",
            "Cv 171": "Z"
        }
        
        self.body_alloys = {
            "CD4MCuN": "0", "316 SS": "1", "Alloy 20": "2", "Monel": "3",
            "Nickel": "4", "Chlorimet 2": "5", "Chlorimet 3": "6",
            "Titanium": "7", "Carbon steel": "8", "304L SS": "9",
            "DV": "C", "Ductile iron": "D", "ZRH": "E", "Zr 705": "G",
            "LCB": "H", "304 SS": "J", "316L SS": "L", "Inconel": "N",
            "CE3MN": "U"
        }
        
        self.plug_alloys = {
            "CD4MCuN": "0", "316 SS": "1", "Alloy 20": "2", "Monel": "3",
            "Nickel": "4", "Chlorimet 2": "5", "Chlorimet 3": "6",
            "Titanium": "7", "Nickel-plated CS": "8", "304L SS": "9",
            "DV": "C", "Nickel-plated DI": "D", "ZRH": "E", "Zr 702": "F",
            "Zr 705": "G", "Nickel-plated LCB": "H", "304 SS": "J",
            "316L SS": "L", "Inconel": "N", "CE3MN": "U"
        }
        
        self.topcap_alloys = {
            "CD4MCuN": "0", "Ductile iron": "D", "316L SS": "L",
            "DC2": "5", "DC3": "6", "Titanium": "7"
        }
        
        self.sleeves = {
            "PTFE": "T", "Durlon II": "D", "Glass Filled": "G",
            "UMPE sleeve w/ PFA diaphragm": "U", "Solid PTFE": "A",
            "Solid Durlon II": "B", "Solid UMPE": "C",
            "TFE/no Zyglo": "X", "TM/no Zyglo": "Z", "UMPE Sleeve": "F"
        }
        
        self.operators = {
            "Std Wrench": "0", "Gear (CI)": "1", "Caustic PF (CI) with CI handwheel": "3",
            "Saginaw": "7", "Bare Stem": "9", "Adjustable T Wrench": "A",
            "Chain wheel operator": "B", "HighHub": "H", "Oval HW": "V"
        }
        
        self.mfg_codes = {"Mfg CVO use only": "M"}
        
        self.adjuster_fasteners = {
            "B840": "1", "B9": "2", "B7": "3", "B7MT": "4", "C20": "5",
            "HC": "6", "I718": "7", "B7M": "9", "B16": "B", "L7": "L",
            "MKH": "M", "B7T": "T", "1840": "G"
        }
        
        self.topcap_fasteners = {
            "B840": "1", "B9": "2", "B7": "3", "B7MT": "4", "C20": "5",
            "HC": "6", "I718": "7", "B7M": "9", "B16": "B", "MKH": "M",
            "B7T": "T", "I825": "I"
        }
        
        self.options = {
            "Drilled & tapped flange": "A", "Block & bleed w/ 1/4DG432 valve": "B",
            "Chlorine service (DNZ)": "C", "Kalrez o-rings": "K",
            "NACE Trim": "N", "Commercial Oxygen Cleaned": "O",
            "Prepared for Phosgene": "P", "Built Dry": "R",
            "Silicone Free": "S", "Vacuum service, 1 micron": "V",
            "None": "W", "Fire sealed - STD": "Z",
            "Dupont Fluoropolymer Material": "@", "Quality Plan": "Q"
        }
        
        self.create_widgets()
        
    def create_widgets(self):
        # Create main canvas with scrollbar
        canvas = tk.Canvas(self.root)
        scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Header
        header = ttk.Label(scrollable_frame, text="Durco G4 Sleeveline Valve Model Number Generator", 
                          font=("Arial", 14, "bold"))
        header.grid(row=0, column=0, columnspan=2, pady=10, padx=10)
        
        row = 1
        
        # Valve Size
        ttk.Label(scrollable_frame, text="Valve Size:", font=("Arial", 10, "bold")).grid(row=row, column=0, sticky="w", padx=10, pady=5)
        self.valve_size = ttk.Combobox(scrollable_frame, values=self.valve_sizes, width=40)
        self.valve_size.grid(row=row, column=1, padx=10, pady=5)
        row += 1
        
        # Product Family
        ttk.Label(scrollable_frame, text="Product Family:", font=("Arial", 10, "bold")).grid(row=row, column=0, sticky="w", padx=10, pady=5)
        self.product_family = ttk.Combobox(scrollable_frame, values=list(self.product_families.keys()), width=40)
        self.product_family.grid(row=row, column=1, padx=10, pady=5)
        row += 1
        
        # Pressure Class
        ttk.Label(scrollable_frame, text="Pressure Class:", font=("Arial", 10, "bold")).grid(row=row, column=0, sticky="w", padx=10, pady=5)
        self.pressure_class = ttk.Combobox(scrollable_frame, values=list(self.pressure_classes.keys()), width=40)
        self.pressure_class.grid(row=row, column=1, padx=10, pady=5)
        row += 1
        
        # End Configuration
        ttk.Label(scrollable_frame, text="End Configuration:", font=("Arial", 10, "bold")).grid(row=row, column=0, sticky="w", padx=10, pady=5)
        self.end_config = ttk.Combobox(scrollable_frame, values=list(self.end_configs.keys()), width=40)
        self.end_config.grid(row=row, column=1, padx=10, pady=5)
        row += 1
        
        # Plug Style
        ttk.Label(scrollable_frame, text="Plug Style:", font=("Arial", 10, "bold")).grid(row=row, column=0, sticky="w", padx=10, pady=5)
        self.plug_style = ttk.Combobox(scrollable_frame, values=list(self.plug_styles.keys()), width=40)
        self.plug_style.grid(row=row, column=1, padx=10, pady=5)
        row += 1
        
        # Body Alloy
        ttk.Label(scrollable_frame, text="Body Alloy:", font=("Arial", 10, "bold")).grid(row=row, column=0, sticky="w", padx=10, pady=5)
        self.body_alloy = ttk.Combobox(scrollable_frame, values=list(self.body_alloys.keys()), width=40)
        self.body_alloy.grid(row=row, column=1, padx=10, pady=5)
        row += 1
        
        # Plug Alloy
        ttk.Label(scrollable_frame, text="Plug Alloy:", font=("Arial", 10, "bold")).grid(row=row, column=0, sticky="w", padx=10, pady=5)
        self.plug_alloy = ttk.Combobox(scrollable_frame, values=list(self.plug_alloys.keys()), width=40)
        self.plug_alloy.grid(row=row, column=1, padx=10, pady=5)
        row += 1
        
        # Top Cap Alloy
        ttk.Label(scrollable_frame, text="Top Cap Alloy:", font=("Arial", 10, "bold")).grid(row=row, column=0, sticky="w", padx=10, pady=5)
        self.topcap_alloy = ttk.Combobox(scrollable_frame, values=list(self.topcap_alloys.keys()), width=40)
        self.topcap_alloy.grid(row=row, column=1, padx=10, pady=5)
        row += 1
        
        # Sleeve/Seats
        ttk.Label(scrollable_frame, text="Sleeve/Seats:", font=("Arial", 10, "bold")).grid(row=row, column=0, sticky="w", padx=10, pady=5)
        self.sleeve = ttk.Combobox(scrollable_frame, values=list(self.sleeves.keys()), width=40)
        self.sleeve.grid(row=row, column=1, padx=10, pady=5)
        row += 1
        
        # Operator
        ttk.Label(scrollable_frame, text="Operator:", font=("Arial", 10, "bold")).grid(row=row, column=0, sticky="w", padx=10, pady=5)
        self.operator = ttk.Combobox(scrollable_frame, values=list(self.operators.keys()), width=40)
        self.operator.grid(row=row, column=1, padx=10, pady=5)
        row += 1
        
        # Manufacturer Code
        ttk.Label(scrollable_frame, text="Manufacturer Code:", font=("Arial", 10, "bold")).grid(row=row, column=0, sticky="w", padx=10, pady=5)
        self.mfg_code = ttk.Combobox(scrollable_frame, values=list(self.mfg_codes.keys()), width=40)
        self.mfg_code.grid(row=row, column=1, padx=10, pady=5)
        row += 1
        
        # Adjuster Fastener
        ttk.Label(scrollable_frame, text="Adjuster Fastener:", font=("Arial", 10, "bold")).grid(row=row, column=0, sticky="w", padx=10, pady=5)
        self.adj_fastener = ttk.Combobox(scrollable_frame, values=list(self.adjuster_fasteners.keys()), width=40)
        self.adj_fastener.grid(row=row, column=1, padx=10, pady=5)
        row += 1
        
        # Top Cap Fastener
        ttk.Label(scrollable_frame, text="Top Cap Fastener:", font=("Arial", 10, "bold")).grid(row=row, column=0, sticky="w", padx=10, pady=5)
        self.topcap_fastener = ttk.Combobox(scrollable_frame, values=list(self.topcap_fasteners.keys()), width=40)
        self.topcap_fastener.grid(row=row, column=1, padx=10, pady=5)
        row += 1
        
        # Options
        ttk.Label(scrollable_frame, text="Options:", font=("Arial", 10, "bold")).grid(row=row, column=0, sticky="w", padx=10, pady=5)
        self.option = ttk.Combobox(scrollable_frame, values=list(self.options.keys()), width=40)
        self.option.grid(row=row, column=1, padx=10, pady=5)
        row += 1
        
        # Generate Button
        generate_btn = ttk.Button(scrollable_frame, text="Generate Model Number", command=self.generate_model)
        generate_btn.grid(row=row, column=0, columnspan=2, pady=20)
        row += 1
        
        # Result Label
        ttk.Label(scrollable_frame, text="Generated Model Number:", font=("Arial", 11, "bold")).grid(row=row, column=0, columnspan=2, pady=5)
        row += 1
        
        self.result_text = tk.Text(scrollable_frame, height=3, width=60, font=("Courier", 12))
        self.result_text.grid(row=row, column=0, columnspan=2, padx=10, pady=5)
        row += 1
        
        # Copy Button
        copy_btn = ttk.Button(scrollable_frame, text="Copy to Clipboard", command=self.copy_to_clipboard)
        copy_btn.grid(row=row, column=0, columnspan=2, pady=10)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
    def generate_model(self):
        try:
            # Build model number according to format
            parts = []
            
            # Valve Size
            if not self.valve_size.get():
                raise ValueError("Please select a valve size")
            parts.append(self.valve_size.get())
            
            # Product Family
            if not self.product_family.get():
                raise ValueError("Please select a product family")
            parts.append(self.product_families[self.product_family.get()])
            
            # Pressure Class + End Config + Plug Style
            if not self.pressure_class.get():
                raise ValueError("Please select a pressure class")
            if not self.end_config.get():
                raise ValueError("Please select an end configuration")
            if not self.plug_style.get():
                raise ValueError("Please select a plug style")
            
            combined = (self.pressure_classes[self.pressure_class.get()] + 
                       self.end_configs[self.end_config.get()] + 
                       self.plug_styles[self.plug_style.get()])
            parts.append(combined)
            
            # Materials (Body + Plug + Top Cap)
            if not self.body_alloy.get():
                raise ValueError("Please select a body alloy")
            if not self.plug_alloy.get():
                raise ValueError("Please select a plug alloy")
            if not self.topcap_alloy.get():
                raise ValueError("Please select a top cap alloy")
            
            materials = (self.body_alloys[self.body_alloy.get()] + 
                        self.plug_alloys[self.plug_alloy.get()] + 
                        self.topcap_alloys[self.topcap_alloy.get()])
            parts.append(materials)
            
            # Sleeve
            if not self.sleeve.get():
                raise ValueError("Please select a sleeve/seat")
            parts.append(self.sleeves[self.sleeve.get()])
            
            # Operator
            if not self.operator.get():
                raise ValueError("Please select an operator")
            parts.append(self.operators[self.operator.get()])
            
            # Manufacturer Code
            if not self.mfg_code.get():
                raise ValueError("Please select a manufacturer code")
            parts.append(self.mfg_codes[self.mfg_code.get()])
            
            # Fasteners (Adjuster + Top Cap)
            if not self.adj_fastener.get():
                raise ValueError("Please select an adjuster fastener")
            if not self.topcap_fastener.get():
                raise ValueError("Please select a top cap fastener")
            
            fasteners = (self.adjuster_fasteners[self.adj_fastener.get()] + 
                        self.topcap_fasteners[self.topcap_fastener.get()])
            parts.append(fasteners)
            
            # Options
            if not self.option.get():
                raise ValueError("Please select an option")
            parts.append(self.options[self.option.get()])
            
            # Build final model number
            model_number = "-".join(parts)
            
            # Display result
            self.result_text.delete(1.0, tk.END)
            self.result_text.insert(1.0, model_number)
            
        except ValueError as e:
            messagebox.showerror("Error", str(e))
        except KeyError as e:
            messagebox.showerror("Error", f"Invalid selection: {e}")
    
    def copy_to_clipboard(self):
        model_number = self.result_text.get(1.0, tk.END).strip()
        if model_number:
            self.root.clipboard_clear()
            self.root.clipboard_append(model_number)
            messagebox.showinfo("Success", "Model number copied to clipboard!")
        else:
            messagebox.showwarning("Warning", "No model number to copy. Please generate one first.")

if __name__ == "__main__":
    root = tk.Tk()
    app = DurcoValveGenerator(root)
    root.mainloop()