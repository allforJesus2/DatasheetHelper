import tkinter as tk
from tkinter import ttk
from tkinter import messagebox

class ScrollableFrame(ttk.Frame):
    """
    A specific frame class that allows the window to be scrollable.
    Essential for long ordering forms.
    """
    def __init__(self, container, *args, **kwargs):
        super().__init__(container, *args, **kwargs)
        
        # Create a canvas and a scrollbar
        self.canvas = tk.Canvas(self)
        self.scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)

        # Configure the scroll region to update when the frame size changes
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        # Draw the frame inside the canvas
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        # Pack the scrollbar and canvas
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")
        
        # Bind mouse wheel for easier scrolling
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")

class ModelBuilderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PGS300/PAS300/PGD300 Model Number Configurator")
        self.root.geometry("900x700")

        # Title Label
        title_lbl = ttk.Label(self.root, text="Pressure Transmitter Model Generator", font=("Arial", 16, "bold"))
        title_lbl.pack(pady=10)

        # Result Display Area (Fixed at top so it's always visible)
        result_frame = ttk.LabelFrame(self.root, text="Generated Model String", padding=10)
        result_frame.pack(fill="x", padx=20, pady=5)
        
        self.result_var = tk.StringVar()
        self.result_entry = ttk.Entry(result_frame, textvariable=self.result_var, font=("Courier", 12, "bold"), state="readonly")
        self.result_entry.pack(fill="x", padx=5)

        # Main Scrollable Area
        self.main_frame = ScrollableFrame(self.root)
        self.main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        self.content = self.main_frame.scrollable_frame

        # --- DATA DEFINITIONS (Based on PDF Pages 27-30) ---
        
        # 1. Mandatory Options (Dropdowns)
        # Format: "Label": {"Display Text": "Code"}
        self.mandatory_data = {
            "Base Model": {
                "PGS300 - Gauge Pressure (0.055%)": "PGS300",
                "PAS300 - Absolute Pressure (0.055%)": "PAS300",
                "PGD300 - Gauge Pressure with Remote Seal (0.055%)": "PGD300"
            },
            "Sensor Span Limits": {
                "0.2 to 6 kPa (C)": "C",
                "4 to 400 mbar (F)": "F",
                "25 to 2500 mbar (L)": "L",
                "0.4 to 40 bar / 16.7-1000 kPa / 6MPa (D)": "D",
                "40 to 4000 kPa (K)": "K",
                "4.14 to 414 bar (R)": "R",
                "6 to 600 bar (T)": "T",
                "105 to 1050 bar (V)": "V"
            },
            "Maximum Working Pressure": {
                "Standard Static (S9)": "S9"
            },
            "Diaphragm Material": {
                "AISI 316L SST (A)": "A",
                "Hastelloy C-276 (H)": "H",
                "AISI 316L SST Gold-plated (G)": "G",
                "H-Shield (S)": "S",
                "Inconel 718 (L)": "L",
                "Diaphragm seal (direct or remote) (R)": "R"
            },
            "Filling Fluid": {
                "Silicon Oil (1)": "1",
                "Inert Fluid (2)": "2",
                "No Filling (N)": "N"
            },
            "Process Connection Material": {
                "AISI 316L SST (A)": "A",
                "Hastelloy C-276 (H)": "H",
                "Inconel 718 (L)": "L",
                "Diaphragm seal (direct or remote) (R)": "R"
            },
            "Process Connection Size": {
                "1/2 in - 14 NPT female (A)": "A",
                "1/2 in NPT male + 1/4 NPT female (B)": "B",
                "DIN EN 837-1 G1/2 in Type B (C)": "C",
                "F250 Autoclave (F)": "F",
                "M20 female (M)": "M",
                "1/4 in - 18 NPT female (H)": "H",
                "Diaphragm seal (direct or remote) (R)": "R"
            },
            "Gasket Material": {
                "None (N)": "N",
                "None (Y)": "Y"
            },
            "Housing Material / Elec. Conn": {
                "Alu Alloy (Barrel) / 1/2-14 NPT (A)": "A",
                "Alu Alloy (Barrel) / M20 x 1.5 (B)": "B",
                "AISI 316L SST (Barrel) / 1/2-14 NPT (S)": "S",
                "AISI 316L SST (Barrel) / M20 x 1.5 (T)": "T"
            },
            "Output": {
                "HART digital & 4-20 mA (H)": "H",
                "Modbus RS 485 (M)": "M",
                "Low-power 1-5 V DC (V)": "V",
                "WirelessHART (W)": "W"
            }
        }

        # 2. Optional Add-ons (Checkboxes)
        # Format: "Category Header": [ ("Description", "Code"), ... ]
        self.optional_data = {
            "Display": [
                ("Integrated digital LCD display (D1)", "D1"),
                ("Integrated digital LCD with backlight (D2)", "D2"),
                ("Integrated digital LCD 4-button (D5)", "D5")
            ],
            "Explosion Protection": [
                ("ATEX Intrinsic Safety Ex ia (HAA)", "HAA"),
                ("ATEX Explosion Proof Ex db_tb (HAD)", "HAD"),
                ("FM Approvals (HUD)", "HUD"),
                ("IECEx Intrinsic Safety (HJA)", "HJA")
            ],
            "Mounting Bracket": [
                ("Pipe/Wall Mount AISI 316L (BA)", "BA"),
                ("Pipe/Wall Mount Carbon Steel (BC)", "BC")
            ],
            "Manifold Assembly": [
                ("Manifold Assembly (A1)", "A1")
            ],
            "Inspection Certificates": [
                ("Calibration 5-point (CF)", "CF"),
                ("Cleanliness stage (CC)", "CC"),
                ("Helium leakage test (CL)", "CL"),
                ("Pressure test (CH)", "CH"),
                ("PMI test on wetted parts (CP)", "CP"),
                ("Material Certificate 3.1 (CM)", "CM")
            ],
            "Manuals": [
                ("German (M1)", "M1"),
                ("Italian (M2)", "M2"),
                ("Spanish (M3)", "M3"),
                ("English (M5)", "M5")
            ],
            "Software Application": [
                ("Calibration Range and/or Tag adjusted to customer specification (SU)", "SU")
            ]
        }

        # Storage for widget references
        self.combo_vars = {}
        self.check_vars = {}

        self.build_ui()
        self.update_model_string() # Initial update

    def build_ui(self):
        # --- Build Mandatory Section ---
        lbl_mandatory = ttk.Label(self.content, text="Base Configuration (Mandatory)", font=("Arial", 12, "bold"))
        lbl_mandatory.grid(row=0, column=0, sticky="w", pady=(0, 10))

        current_row = 1
        
        for category, options in self.mandatory_data.items():
            # Label
            lbl = ttk.Label(self.content, text=category + ":")
            lbl.grid(row=current_row, column=0, sticky="w", padx=10, pady=5)
            
            # Variable
            var = tk.StringVar()
            self.combo_vars[category] = var
            
            # Combobox
            combo = ttk.Combobox(self.content, textvariable=var, state="readonly", width=50)
            combo['values'] = list(options.keys())
            combo.current(0) # Select first option by default
            combo.grid(row=current_row, column=1, sticky="w", padx=10, pady=5)
            
            # Bind event to update string
            combo.bind("<<ComboboxSelected>>", self.update_model_string)
            
            current_row += 1

        ttk.Separator(self.content, orient='horizontal').grid(row=current_row, column=0, columnspan=2, sticky="ew", pady=20)
        current_row += 1

        # --- Build Optional Section ---
        lbl_optional = ttk.Label(self.content, text="Additional Options (Optional)", font=("Arial", 12, "bold"))
        lbl_optional.grid(row=current_row, column=0, sticky="w", pady=(0, 10))
        current_row += 1

        for category, items in self.optional_data.items():
            # Category Sub-header
            sub_lbl = ttk.Label(self.content, text=category, font=("Arial", 10, "italic"))
            sub_lbl.grid(row=current_row, column=0, sticky="w", padx=10, pady=(10, 5))
            current_row += 1
            
            # Create a frame for checkboxes to organize them nicely
            chk_frame = ttk.Frame(self.content)
            chk_frame.grid(row=current_row, column=0, columnspan=2, sticky="w", padx=20)
            
            # Grid columns for checkboxes inside the frame
            col_idx = 0
            row_idx = 0
            
            for desc, code in items:
                var = tk.BooleanVar()
                self.check_vars[(category, desc)] = (var, code)
                
                chk = ttk.Checkbutton(chk_frame, text=desc, variable=var, command=self.update_model_string)
                chk.grid(row=row_idx, column=col_idx, sticky="w", padx=10, pady=2)
                
                # Wrap checkboxes to next row after 2 columns
                col_idx += 1
                if col_idx > 1:
                    col_idx = 0
                    row_idx += 1

            current_row += 1

    def update_model_string(self, event=None):
        """Generates the full model code based on selections."""
        model_parts = []
        
        # 1. Gather Mandatory Codes
        for category, options_dict in self.mandatory_data.items():
            selection_text = self.combo_vars[category].get()
            # Retrieve the code from the dictionary
            code = options_dict.get(selection_text, "")
            model_parts.append(code)

        # 2. Gather Optional Codes
        optional_codes = []
        for key, (var, code) in self.check_vars.items():
            if var.get():
                optional_codes.append(code)

        # 3. Construct Final String
        # Base model parts are usually joined directly (e.g., PGS300C...)
        # PGD300 models use "-" separator before optional codes (e.g., PGD300DS9R1RRYAH-D1HUDSU)
        # PGS300/PAS300 models use space separator
        
        base_str = "".join(model_parts)
        
        # Check if it's a PGD300 model - use "-" separator
        if base_str.startswith("PGD300"):
            if optional_codes:
                full_str = f"{base_str}-" + "".join(optional_codes)
            else:
                full_str = base_str
        else:
            # PGS300/PAS300 models use space separator
            if optional_codes:
                full_str = f"{base_str} " + " ".join(optional_codes)
            else:
                full_str = base_str

        self.result_var.set(full_str)

if __name__ == "__main__":
    root = tk.Tk()
    # Set a theme for better visuals
    style = ttk.Style()
    style.theme_use('clam') 
    app = ModelBuilderApp(root)
    root.mainloop()