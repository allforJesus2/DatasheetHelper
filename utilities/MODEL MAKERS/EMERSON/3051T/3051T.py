import tkinter as tk
from tkinter import ttk, messagebox

class Rosemount3051TConfigurator:
    def __init__(self, root):
        self.root = root
        self.root.title("Rosemount 3051T In-Line Pressure Transmitter Configurator")
        self.root.geometry("950x750")
        
        # Model number components storage
        self.selections = {}
        
        # Create main container with scrollbar
        main_frame = ttk.Frame(root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Create canvas and scrollbar
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
        title = ttk.Label(scrollable_frame, text="Rosemount 3051T In-Line Pressure Transmitter", 
                         font=("Arial", 16, "bold"))
        title.grid(row=0, column=0, columnspan=3, pady=10, sticky="w")
        
        # Instructions
        instructions = ttk.Label(scrollable_frame, 
                                text="Select required options below to generate a complete model number",
                                font=("Arial", 10))
        instructions.grid(row=1, column=0, columnspan=3, pady=5, sticky="w")
        
        row = 2
        
        # Model (Base)
        row = self.add_dropdown(scrollable_frame, row, "Model", "model",
            {"3051T": "In-Line Pressure Transmitter"},
            "Base model for in-line pressure measurement")
        
        # Pressure Type
        row = self.add_dropdown(scrollable_frame, row, "Pressure Type", "pressure_type",
            {
                "G": "Gage Pressure",
                "A": "Absolute Pressure (also allows wireless)"
            },
            "Type of pressure measurement")
        
        # Pressure Range
        row = self.add_dropdown(scrollable_frame, row, "Pressure Range", "range",
            {
                "0": "–5 to 5 psi (Gage only)",
                "1": "–14.7 to 30 psi (Gage) / 0 to 30 psia (Absolute)",
                "2": "–14.7 to 150 psi (Gage) / 0 to 150 psia (Absolute)",
                "3": "–14.7 to 800 psi (Gage) / 0 to 800 psia (Absolute)",
                "4": "–14.7 to 4000 psi (Gage) / 0 to 4000 psia (Absolute)",
                "5": "–14.7 to 10000 psi (Gage) / 0 to 10000 psia (Absolute)",
                "6": "–14.7 to 20000 psi (Gage) / 0 to 20000 psia (Absolute)"
            },
            "Pressure measurement range")
        
        # Transmitter Output
        row = self.add_dropdown(scrollable_frame, row, "Transmitter Output", "output",
            {
                "A": "4–20 mA with HART Protocol",
                "F": "FOUNDATION Fieldbus Protocol",
                "W": "PROFIBUS PA Protocol",
                "X": "Wireless (requires wireless options)",
                "M": "Low-power 1–5 Vdc with HART Protocol"
            },
            "Communication protocol and output signal")
        
        # Process Connection Style
        row = self.add_dropdown(scrollable_frame, row, "Process Connection", "connection",
            {
                "2B": "½-14 NPT female (range 0-5 only)",
                "2C": "G½ A EN837-1 male (range 0-4 only)",
                "2F": "Coned/threaded autoclave F-250-C (range 5-6 only)",
                "61": "Non-threaded instrument flange (range 1-4 only)"
            },
            "Style and size of process connection")
        
        # Isolating Diaphragm Material
        row = self.add_dropdown(scrollable_frame, row, "Isolating Diaphragm", "diaphragm",
            {
                "2": "316L stainless steel",
                "3": "Alloy C-276",
                "7": "Gold-plated 316 stainless steel",
                "8": "Gold-plated alloy C-276"
            },
            "Diaphragm and process wetted material")
        
        # Sensor Fill Fluid
        row = self.add_dropdown(scrollable_frame, row, "Sensor Fill Fluid", "fill_fluid",
            {
                "1": "Silicone",
                "2": "Inert (not for wireless)"
            },
            "Internal sensor fill fluid type")
        
        # Housing Material
        row = self.add_dropdown(scrollable_frame, row, "Housing Material", "housing",
            {
                "A": "Aluminum, ½-14 NPT conduit entry",
                "B": "Aluminum, M20 x 1.5 conduit entry",
                "E": "Aluminum ultra low copper, ½-14 NPT",
                "F": "Aluminum ultra low copper, M20 x 1.5",
                "J": "Stainless steel, ½-14 NPT",
                "K": "Stainless steel, M20 x 1.5",
                "P": "Engineered polymer (wireless only, no conduit)",
                "D": "Aluminum, G½ (with ½ NPT adapter)",
                "M": "Stainless steel, G½ (with ½ NPT adapter)"
            },
            "Transmitter housing material and conduit entry")
        
        # Separator for Optional Codes
        ttk.Separator(scrollable_frame, orient='horizontal').grid(
            row=row, column=0, columnspan=3, sticky='ew', pady=10)
        row += 1
        
        # Optional Codes Section Header
        optional_header = ttk.Label(scrollable_frame, 
                                   text="Optional Codes (Select if needed)", 
                                   font=("Arial", 12, "bold"), 
                                   foreground="blue")
        optional_header.grid(row=row, column=0, columnspan=3, sticky="w", pady=5)
        row += 1
        
        # Wireless Options (if applicable)
        row = self.add_dropdown(scrollable_frame, row, "Wireless Transmit Rate (Optional)", "wireless_rate",
            {
                "": "None (not wireless)",
                "WA3": "User configurable, 2.4 GHz WirelessHART"
            },
            "Wireless communication rate (requires output code X)")
        
        row = self.add_dropdown(scrollable_frame, row, "Wireless Antenna (Optional)", "wireless_antenna",
            {
                "": "None (not wireless)",
                "WP5": "Internal antenna with SmartPower compatibility"
            },
            "Wireless antenna configuration (requires output code X)")
        
        # Bluetooth
        row = self.add_dropdown(scrollable_frame, row, "Bluetooth Access (Optional)", "bluetooth",
            {
                "": "None",
                "BLE": "Bluetooth configuration and maintenance"
            },
            "Wireless configuration via Bluetooth (requires M6 display)")
        
        # Extended Warranty
        row = self.add_dropdown(scrollable_frame, row, "Extended Warranty (Optional)", "warranty",
            {
                "": "Standard warranty",
                "WR3": "3-year limited warranty",
                "WR5": "5-year limited warranty"
            },
            "Extended product warranty coverage")
        
        # Plantweb Control
        row = self.add_dropdown(scrollable_frame, row, "Plantweb Control (Optional)", "plantweb_control",
            {
                "": "None",
                "A01": "FOUNDATION Fieldbus control function block suite"
            },
            "Advanced control functionality")
        
        # Plantweb Diagnostics
        row = self.add_dropdown(scrollable_frame, row, "Plantweb Diagnostics (Optional)", "diagnostics",
            {
                "": "None",
                "DA0": "Loop Integrity Diagnostic (HART only)",
                "DA1": "Loop Integrity + Plugged Impulse Line (HART/Low Power)",
                "D01": "FOUNDATION Fieldbus Diagnostics Suite"
            },
            "Advanced diagnostic capabilities")
        
        # Integral Assembly
        row = self.add_dropdown(scrollable_frame, row, "Integral Assembly (Optional)", "integral",
            {
                "": "None",
                "S5": "Assemble to Rosemount 306 Integral Manifold",
                "S1": "Assemble to one Rosemount seal"
            },
            "Pre-assembled components")
        
        # Mounting Bracket
        row = self.add_dropdown(scrollable_frame, row, "Mounting Bracket (Optional)", "bracket",
            {
                "": "None",
                "B4": "2-in. pipe or panel mounting, stainless steel",
                "BE": "316SST B4 bracket with 316SST bolts"
            },
            "Mounting hardware")
        
        # Product Certifications
        row = self.add_dropdown(scrollable_frame, row, "Product Certification (Optional)", "certification",
            {
                "": "None (not for hazardous areas)",
                "E1": "ATEX Flameproof",
                "I1": "ATEX Intrinsic Safety",
                "IA": "ATEX FISCO Intrinsic Safety",
                "N1": "ATEX Type n",
                "K8": "ATEX Combination (E8, I1, N1)",
                "E5": "USA Explosion-proof, Dust Ignition-proof",
                "I5": "USA Intrinsically Safe, Nonincendive",
                "K5": "USA Combination (Explosion-proof, IS, Div 2)",
                "C6": "Canada Combination (Explosion-proof, IS, Div 2)",
                "KB": "USA+Canada Combination",
                "KS": "USA+Canada+IECEx+ATEX Full Combination",
                "E7": "IECEx Flameproof",
                "I7": "IECEx Intrinsic Safety",
                "K7": "IECEx Combination"
            },
            "Hazardous area approvals")
        
        # Display Options
        row = self.add_dropdown(scrollable_frame, row, "Display Option (Optional)", "display",
            {
                "": "None",
                "M6": "Graphical LCD display (HART only)",
                "M5": "LCD display",
                "M4": "LCD display with LOI (HART/Low Power/PROFIBUS)"
            },
            "Local display interface")
        
        # Configuration Buttons
        row = self.add_dropdown(scrollable_frame, row, "Configuration Buttons (Optional)", "config_buttons",
            {
                "": "None",
                "D1": "Quick service buttons (requires M6)",
                "D4": "Analog zero and span (HART only)",
                "DZ": "Digital zero trim (HART/Wireless only)"
            },
            "Physical configuration buttons")
        
        # Special Approvals
        row = self.add_dropdown(scrollable_frame, row, "Special Approvals (Optional)", "special_approval",
            {
                "": "None",
                "DW": "NSF drinking water approval",
                "SBS": "American Bureau of Shipping",
                "SDN": "Det Norske Veritas",
                "SLL": "Lloyds Register"
            },
            "Industry-specific certifications")
        
        # Quality Options
        row = self.add_dropdown(scrollable_frame, row, "Quality/Safety (Optional)", "quality",
            {
                "": "None",
                "Q4": "Calibration certificate",
                "QP": "Calibration certificate + tamper evident seal",
                "Q8": "Material traceability (EN 10204 3.1.B)",
                "Q76": "PMI verification and certificate",
                "QT": "Safety certified to IEC 61508 (HART only)",
                "T9": "Enhanced SIS proof testing (HART only)"
            },
            "Quality and safety documentation")
        
        # NACE Certificate
        row = self.add_dropdown(scrollable_frame, row, "NACE Certificate (Optional)", "nace",
            {
                "": "None",
                "Q15": "NACE MR0175/ISO 15156 (Sour oil field)",
                "Q25": "NACE MR0103 (Sour refining)"
            },
            "Sour service environment compliance")
        
        # Transient Protection
        row = self.add_dropdown(scrollable_frame, row, "Transient Protection (Optional)", "transient",
            {
                "": "None",
                "T1": "Transient protection terminal block (not for wireless)"
            },
            "Electrical transient protection")
        
        # Other Hardware Options
        row = self.add_dropdown(scrollable_frame, row, "Additional Hardware (Optional)", "hardware",
            {
                "": "None",
                "DO": "316 SST conduit plug (not for wireless)",
                "V5": "External ground screw assembly (not for wireless)",
                "WSM": "Wireless stainless steel sensor module"
            },
            "Additional hardware components")
        
        # Software/Configuration
        row = self.add_dropdown(scrollable_frame, row, "Software Config (Optional)", "software",
            {
                "": "None",
                "C1": "Custom software configuration",
                "RK": "Enhanced software (expanded alerts/logging)"
            },
            "Software configuration options")
        
        # Alarm Levels
        row = self.add_dropdown(scrollable_frame, row, "Alarm Levels (Optional)", "alarm",
            {
                "": "Standard",
                "C4": "NAMUR NE 43, alarm high (HART only)",
                "CN": "NAMUR NE 43, alarm low (HART only)",
                "CT": "Rosemount standard low alarm (HART only)",
                "CR": "Custom high alarm (requires C1, HART only)",
                "CS": "Custom low alarm (requires C1, HART only)"
            },
            "Alarm signal level configuration")
        
        # Process Testing/Cleaning
        row = self.add_dropdown(scrollable_frame, row, "Testing/Cleaning (Optional)", "testing",
            {
                "": "None",
                "P1": "Hydrostatic testing with certificate",
                "P2": "Cleaning for special service",
                "P3": "Cleaning for < 1 ppm chlorine/fluoride"
            },
            "Factory testing and cleaning services")
        
        # Electrical Connectors
        row = self.add_dropdown(scrollable_frame, row, "Electrical Connector (Optional)", "connector",
            {
                "": "Standard conduit entry",
                "GB": "ATEX cable gland and plug, M20 x 1.5",
                "GE": "M12, 4-pin male connector (eurofast)",
                "GM": "A size mini, 4-pin male connector (minifast)"
            },
            "Alternative electrical connections")
        
        # Cold Temperature
        row = self.add_dropdown(scrollable_frame, row, "Cold Temperature (Optional)", "cold_temp",
            {
                "": "Standard temperature range",
                "BR5": "–58 °F (–50 °C) operation (range 1-5, silicone)",
                "BR6": "–76 °F (–60 °C) operation (range 1-5, silicone)"
            },
            "Extended low temperature operation")
        
        # Other Options
        row = self.add_dropdown(scrollable_frame, row, "Other Options (Optional)", "other",
            {
                "": "None",
                "Y2": "316 SST nameplate, top tag, wire-on tag",
                "C5": "Measurement Canada Accuracy Approval",
                "Q16": "Surface finish certification for sanitary seals",
                "QZ": "Remote seal system performance report"
            },
            "Additional special options")
        
        # Separator
        ttk.Separator(scrollable_frame, orient='horizontal').grid(
            row=row, column=0, columnspan=3, sticky='ew', pady=10)
        row += 1
        
        # Generate Button
        generate_btn = ttk.Button(scrollable_frame, text="Generate Model Number", 
                                 command=self.generate_model, style="Accent.TButton")
        generate_btn.grid(row=row, column=0, columnspan=3, pady=10)
        row += 1
        
        # Result Display
        result_label = ttk.Label(scrollable_frame, text="Generated Model Number:", 
                               font=("Arial", 12, "bold"))
        result_label.grid(row=row, column=0, columnspan=3, sticky="w", pady=(10,5))
        row += 1
        
        self.result_text = tk.Text(scrollable_frame, height=6, width=90, 
                                   font=("Courier", 10), wrap=tk.WORD)
        self.result_text.grid(row=row, column=0, columnspan=3, pady=5, sticky="ew")
        row += 1
        
        # Copy Button
        copy_btn = ttk.Button(scrollable_frame, text="Copy to Clipboard", 
                             command=self.copy_to_clipboard)
        copy_btn.grid(row=row, column=0, columnspan=3, pady=5)
        row += 1
        
        # Clear Button
        clear_btn = ttk.Button(scrollable_frame, text="Clear All Selections", 
                              command=self.clear_all)
        clear_btn.grid(row=row, column=0, columnspan=3, pady=5)
        
        # Pack canvas and scrollbar
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Configure style
        style = ttk.Style()
        style.configure("Accent.TButton", font=("Arial", 11, "bold"))
        
    def add_dropdown(self, parent, row, label_text, key, options, description):
        """Add a labeled dropdown with description"""
        # Label
        label = ttk.Label(parent, text=f"{label_text}:", font=("Arial", 10, "bold"))
        label.grid(row=row, column=0, sticky="w", pady=5, padx=(0,10))
        
        # Dropdown
        var = tk.StringVar()
        dropdown = ttk.Combobox(parent, textvariable=var, width=55, state="readonly")
        dropdown['values'] = [f"{code} - {desc}" for code, desc in options.items()]
        dropdown.grid(row=row, column=1, sticky="w", pady=5)
        
        # Store reference
        self.selections[key] = {"var": var, "options": options, "widget": dropdown}
        
        # Description
        desc_label = ttk.Label(parent, text=description, font=("Arial", 8), 
                              foreground="gray")
        desc_label.grid(row=row+1, column=1, sticky="w", pady=(0,5))
        
        return row + 2
    
    def generate_model(self):
        """Generate the complete model number"""
        model_parts = []
        missing = []
        
        # Required fields
        required_order = [
            "model", "pressure_type", "range", "output", "connection",
            "diaphragm", "fill_fluid", "housing"
        ]
        
        # Optional fields
        optional_order = [
            "wireless_rate", "wireless_antenna", "bluetooth", "warranty",
            "plantweb_control", "diagnostics", "integral", "bracket",
            "certification", "display", "config_buttons", "special_approval",
            "quality", "nace", "transient", "hardware", "software", "alarm",
            "testing", "connector", "cold_temp", "other"
        ]
        
        # Process required fields
        for key in required_order:
            value = self.selections[key]["var"].get()
            if value:
                code = value.split(" - ")[0]
                model_parts.append(code)
            else:
                field_names = {
                    "model": "Model",
                    "pressure_type": "Pressure Type",
                    "range": "Pressure Range",
                    "output": "Transmitter Output",
                    "connection": "Process Connection",
                    "diaphragm": "Isolating Diaphragm",
                    "fill_fluid": "Sensor Fill Fluid",
                    "housing": "Housing Material"
                }
                missing.append(field_names.get(key, key))
        
        if missing:
            messagebox.showwarning("Incomplete Selection", 
                                  f"Please select values for:\n" + "\n".join(f"• {m}" for m in missing))
            return
        
        # Validate configuration
        validation_errors = self.validate_configuration(model_parts)
        if validation_errors:
            messagebox.showerror("Invalid Configuration", 
                               "The following issues were found:\n\n" + 
                               "\n\n".join(f"• {err}" for err in validation_errors))
            return
        
        # Process optional fields
        optional_codes = []
        for key in optional_order:
            value = self.selections[key]["var"].get()
            if value and value != "":
                code = value.split(" - ")[0]
                if code and code != "":
                    optional_codes.append(code)
        
        # Validate optional codes
        optional_errors = self.validate_optional_codes(model_parts, optional_codes)
        if optional_errors:
            messagebox.showerror("Invalid Optional Configuration", 
                               "The following issues were found:\n\n" + 
                               "\n\n".join(f"• {err}" for err in optional_errors))
            return
        
        # Build model number
        model_number = "-".join(model_parts)
        if optional_codes:
            model_number += "-" + "-".join(optional_codes)
        
        # Display result
        self.result_text.delete(1.0, tk.END)
        self.result_text.insert(1.0, f"Complete Model Number:\n\n{model_number}\n\n")
        
        summary = f"Base Configuration: {len(model_parts)} codes\n"
        if optional_codes:
            summary += f"Optional Codes: {len(optional_codes)} selected\n"
            summary += f"Options: {', '.join(optional_codes)}\n\n"
        else:
            summary += "Optional Codes: None selected\n\n"
        
        self.result_text.insert(tk.END, summary)
        self.result_text.insert(tk.END, "Configuration complete. Review selections before ordering.")
    
    def validate_configuration(self, parts):
        """Validate base configuration"""
        errors = []
        
        pressure_type = parts[1] if len(parts) > 1 else ""
        range_code = parts[2] if len(parts) > 2 else ""
        output = parts[3] if len(parts) > 3 else ""
        connection = parts[4] if len(parts) > 4 else ""
        diaphragm = parts[5] if len(parts) > 5 else ""
        fill_fluid = parts[6] if len(parts) > 6 else ""
        housing = parts[7] if len(parts) > 7 else ""
        
        # Rule 1: Wireless only with Absolute pressure
        if output == "X" and pressure_type != "A":
            errors.append("Wireless output (X) only available with Absolute pressure type (A)")
        
        # Rule 2: Wireless only with ranges 1-5
        if output == "X" and range_code not in ["1", "2", "3", "4", "5"]:
            errors.append("Wireless output (X) only available with pressure ranges 1-5")
        
        # Rule 3: Wireless requires engineered polymer housing
        if output == "X" and housing != "P":
            errors.append("Wireless output (X) requires engineered polymer housing (P)")
        
        # Rule 4: Engineered polymer only for wireless
        if housing == "P" and output != "X":
            errors.append("Engineered polymer housing (P) only available with wireless output (X)")
        
        # Rule 5: Engineered polymer only with gauge ranges 1-4
        if housing == "P" and pressure_type == "G" and range_code not in ["1", "2", "3", "4"]:
            errors.append("Engineered polymer housing (P) only available with gauge ranges 1-4")
        
        # Rule 6: Range 6 not available with PROFIBUS PA or Low Power
        if range_code == "6" and output in ["W", "M"]:
            errors.append("Range 6 (20000 psi) not available with PROFIBUS PA (W) or Low Power (M)")
        
        # Rule 7: Range 6 not available with inert fill fluid
        if range_code == "6" and fill_fluid == "2":
            errors.append("Range 6 (20000 psi) not available with inert sensor fill fluid (2)")
        
        # Rule 8: Inert fill not available with wireless
        if fill_fluid == "2" and output == "X":
            errors.append("Inert sensor fill fluid (2) not available with wireless output (X)")
        
        # Rule 9: Connection 2B only for range 0-5
        if connection == "2B" and range_code not in ["0", "1", "2", "3", "4", "5"]:
            errors.append("½-14 NPT female connection (2B) only available for ranges 0-5")
        
        # Rule 10: Connection 2C only for range 0-4
        if connection == "2C" and range_code not in ["0", "1", "2", "3", "4"]:
            errors.append("G½ A EN837-1 connection (2C) only available for ranges 0-4")
        
        # Rule 11: Connection 2F only for range 5-6
        if connection == "2F" and range_code not in ["5", "6"]:
            errors.append("Autoclave connection (2F) only available for ranges 5-6")
        
        # Rule 12: Connection 61 only for range 1-4
        if connection == "61" and range_code not in ["1", "2", "3", "4"]:
            errors.append("Non-threaded instrument flange (61) only available for ranges 1-4")
        
        # Rule 13: Connection 61 requires 316L SST diaphragm
        if connection == "61" and diaphragm != "2":
            errors.append("Non-threaded instrument flange (61) requires 316L SST diaphragm (2)")
        
        # Rule 14: Connection 61 not available with wireless
        if connection == "61" and output == "X":
            errors.append("Non-threaded instrument flange (61) not available with wireless (X)")
        
        # Rule 15: Connection 2C not with wireless for absolute or C-276
        if connection == "2C" and output == "X" and (pressure_type == "A" or diaphragm == "3"):
            errors.append("G½ connection (2C) not available with wireless (X) for absolute pressure or C-276 diaphragm")
        
        # Rule 16: Connection 2F not with wireless for range 5
        if connection == "2F" and output == "X" and range_code == "5":
            errors.append("Autoclave connection (2F) not available with wireless (X) for range 5")
        
        return errors
    
    def validate_optional_codes(self, base_parts, optional_codes):
        """Validate optional codes"""
        errors = []
        
        output = base_parts[3] if len(base_parts) > 3 else ""
        range_code = base_parts[2] if len(base_parts) > 2 else ""
        connection = base_parts[4] if len(base_parts) > 4 else ""
        diaphragm = base_parts[5] if len(base_parts) > 5 else ""
        fill_fluid = base_parts[6] if len(base_parts) > 6 else ""
        housing = base_parts[7] if len(base_parts) > 7 else ""
        
        # Wireless validation
        wireless_codes = [c for c in optional_codes if c in ["WA3", "WP5"]]
        if wireless_codes and output != "X":
            errors.append(f"Wireless options ({', '.join(wireless_codes)}) require wireless output (X)")
        
        if output == "X" and not any(c in optional_codes for c in ["WA3", "WP5"]):
            errors.append("Wireless output (X) requires wireless options WA3 and WP5")
        
        # Bluetooth requires M6 display
        if "BLE" in optional_codes and "M6" not in optional_codes:
            errors.append("Bluetooth (BLE) requires Graphical LCD Display (M6)")
        
        # Display restrictions
        if "M6" in optional_codes and output != "A":
            errors.append("Graphical LCD Display (M6) only available with HART output (A)")
        
        if "M4" in optional_codes and output not in ["A", "M", "W"]:
            errors.append("LCD with LOI (M4) only available with HART (A), Low Power (M), or PROFIBUS (W)")
        
        # Configuration button restrictions
        if "D1" in optional_codes and "M6" not in optional_codes:
            errors.append("Quick service buttons (D1) require Graphical LCD Display (M6)")
        
        if "D4" in optional_codes and output != "A":
            errors.append("Analog zero/span buttons (D4) only available with HART output (A)")
        
        if "DZ" in optional_codes and output not in ["A", "X"]:
            errors.append("Digital zero trim (DZ) only available with HART (A) or Wireless (X)")
        
        # Diagnostic restrictions
        if "DA0" in optional_codes and output != "A":
            errors.append("Loop Integrity Diagnostic (DA0) only available with HART output (A)")
        
        if "DA1" in optional_codes and output not in ["A", "M"]:
            errors.append("Loop Integrity + Plugged Line (DA1) only available with HART (A) or Low Power (M)")
        
        if "D01" in optional_codes and output != "F":
            errors.append("FOUNDATION Fieldbus Diagnostics (D01) only available with Fieldbus output (F)")
        
        if "A01" in optional_codes and output != "F":
            errors.append("FOUNDATION Fieldbus control blocks (A01) only available with Fieldbus output (F)")
        
        # Integral assembly restrictions
        if "S5" in optional_codes and connection == "2C":
            errors.append("Manifold assembly (S5) not available with G½ connection (2C)")
        
        if "S1" in optional_codes and connection == "2C":
            errors.append("Seal assembly (S1) not available with G½ connection (2C)")
        
        if "S5" in optional_codes and diaphragm == "3":
            errors.append("Manifold assembly (S5) not available with Alloy C-276 diaphragm (3)")
        
        if "S5" in optional_codes and range_code == "6":
            errors.append("Manifold assembly (S5) not available with range 6 (20000 psi)")
        
        # Product certification restrictions
        if "E4" in optional_codes and output not in ["A", "F", "W"]:
            errors.append("Japan certification (E4) only with HART (A), Fieldbus (F), or PROFIBUS (W)")
        
        if "E4" in optional_codes and housing != "D":
            errors.append("Japan certification (E4) requires aluminum housing with G½ entry (D)")
        
        if "W" in output and "M4" not in optional_codes:
            cert_codes = [c for c in optional_codes if c in ["E4", "EM", "EP", "I6", "IM", "KD", "KL", "KM", "KP", "KS", "N3"]]
            if cert_codes:
                errors.append(f"PROFIBUS (W) with certifications {', '.join(cert_codes)} requires LOI display (M4)")
        
        if "M" in output:
            allowed_certs = ["C6", "E2", "E5", "I5", "K5", "KB", "EM", "IM", "KM", "EP", "E8"]
            cert_codes = [c for c in optional_codes if c in ["E1", "I1", "IA", "N1", "K8", "E4", "E5", "I5", "K5", "E6", "I6", "C6", "K6", "E7", "I7", "N7", "K7", "IG", "E2", "I2", "IB", "K2", "E3", "I3", "EM", "IM", "KM", "KB", "KD", "KL", "KS", "EP", "IP", "KP"]]
            invalid_certs = [c for c in cert_codes if c not in allowed_certs]
            if invalid_certs:
                errors.append(f"Low Power output (M) not compatible with certifications: {', '.join(invalid_certs)}")
        
        # Wireless certification restrictions
        if output == "X" and "I1" in optional_codes:
            errors.append("Wireless (X) dust approval not applicable with ATEX IS (I1) - see wireless manual")
        
        if output == "X" and "I5" in optional_codes:
            errors.append("Wireless (X) nonincendive certification not provided with USA IS (I5)")
        
        if output == "X" and "KL" not in optional_codes:
            cert_codes = [c for c in optional_codes if c not in ["KL"] and c in ["E1", "I1", "IA", "N1", "K8", "E5", "I5", "K5", "C6", "KB", "KS"]]
            if cert_codes:
                errors.append(f"Wireless (X) only supports certification KL for standard wireless, not: {', '.join(cert_codes)}")
        
        # Drinking water restrictions
        if "DW" in optional_codes:
            if diaphragm == "3":
                errors.append("Drinking water approval (DW) not available with Alloy C-276 diaphragm (3)")
            if "S5" in optional_codes:
                errors.append("Drinking water approval (DW) not available with manifold assembly (S5)")
            if "S1" in optional_codes:
                errors.append("Drinking water approval (DW) not available with seal assembly (S1)")
            if "Q16" in optional_codes:
                errors.append("Drinking water approval (DW) not available with surface finish cert (Q16)")
        
        # Shipboard approval restrictions
        shipboard_codes = [c for c in optional_codes if c in ["SBS", "SBV", "SDN", "SLL"]]
        if shipboard_codes:
            if output == "X":
                errors.append(f"Shipboard approvals ({', '.join(shipboard_codes)}) not available with wireless (X)")
            
            if any(c in ["SBV", "SLL"] for c in shipboard_codes):
                valid_certs = ["E7", "E8", "I1", "I7", "IA", "K7", "K8", "KD", "N1", "N7"]
                has_valid = any(c in optional_codes for c in valid_certs)
                if not has_valid:
                    errors.append(f"Bureau Veritas/Lloyds (SBV/SLL) require specific certifications: {', '.join(valid_certs)}")
        
        # Quality certification restrictions
        if "C5" in optional_codes and output != "A":
            errors.append("Custody transfer (C5) only available with HART output (A)")
        
        if "QT" in optional_codes and output != "A":
            errors.append("Safety certification (QT) only available with HART output (A)")
        
        if "T9" in optional_codes and output != "A":
            errors.append("Enhanced SIS testing (T9) only available with HART output (A)")
        
        # Alarm level restrictions
        alarm_codes = [c for c in optional_codes if c in ["C4", "CN", "CR", "CS", "CT"]]
        if alarm_codes and output != "A":
            errors.append(f"Alarm levels ({', '.join(alarm_codes)}) only available with HART output (A)")
        
        if any(c in ["CR", "CS"] for c in alarm_codes) and "C1" not in optional_codes:
            errors.append("Custom alarm levels (CR/CS) require custom software configuration (C1)")
        
        # Hardware restrictions
        hardware_wireless = [c for c in optional_codes if c in ["T1", "DO", "V5"]]
        if hardware_wireless and output == "X":
            errors.append(f"Hardware options ({', '.join(hardware_wireless)}) not available with wireless (X)")
        
        if "T1" in optional_codes and "V5" in optional_codes:
            errors.append("External ground screw (V5) not needed with transient protection (T1) - already included")
        
        fisco_certs = [c for c in optional_codes if c in ["IA", "IB", "IG"]]
        if fisco_certs and "T1" in optional_codes:
            errors.append(f"Transient protection (T1) not needed with FISCO certifications ({', '.join(fisco_certs)})")
        
        # Connector restrictions
        connector_codes = [c for c in optional_codes if c in ["GB", "GE", "GM"]]
        if connector_codes and output == "X":
            errors.append(f"Electrical connectors ({', '.join(connector_codes)}) not available with wireless (X)")
        
        # Testing restrictions
        if "P1" in optional_codes and range_code == "0":
            errors.append("Hydrostatic testing (P1) not available with range 0 (±5 psi)")
        
        if any(c in ["P2", "P3"] for c in optional_codes) and "S5" in optional_codes:
            errors.append("Special cleaning (P2/P3) not valid with manifold assembly (S5)")
        
        # Cold temperature restrictions
        cold_codes = [c for c in optional_codes if c in ["BR5", "BR6"]]
        if cold_codes:
            if range_code not in ["1", "2", "3", "4", "5"]:
                errors.append(f"Cold temperature ({', '.join(cold_codes)}) only available with ranges 1-5")
            if output not in ["A", "F"]:
                errors.append(f"Cold temperature ({', '.join(cold_codes)}) only with HART (A) or Fieldbus (F)")
            if fill_fluid != "1":
                errors.append(f"Cold temperature ({', '.join(cold_codes)}) requires silicone fill fluid (1)")
            if connection == "61":
                errors.append(f"Cold temperature ({', '.join(cold_codes)}) not available with instrument flange (61)")
            if "S1" in optional_codes:
                errors.append(f"Cold temperature ({', '.join(cold_codes)}) not available with seal assembly (S1)")
        
        if "BR5" in optional_codes:
            valid_br5_certs = ["C6", "E2", "E5", "E6", "E7", "EM", "EP", "I2", "I5", "I6", "I7", "IM", "IP", "K2", "K5", "K7", "KB", "KM", "KP"]
            cert_codes = [c for c in optional_codes if c in ["E1", "I1", "IA", "N1", "K8", "E4", "E5", "I5", "K5", "E6", "I6", "C6", "K6", "E7", "I7", "N7", "K7", "IG", "E2", "I2", "IB", "K2", "E3", "I3", "EM", "IM", "KM", "KB", "KD", "KL", "KS", "EP", "IP", "KP"]]
            invalid_certs = [c for c in cert_codes if c not in valid_br5_certs]
            if invalid_certs:
                errors.append(f"BR5 cold temp not compatible with certifications: {', '.join(invalid_certs)}")
        
        if "BR6" in optional_codes:
            valid_br6_certs = ["E2", "E7", "EM", "I2", "I6", "I7", "IM", "IP", "K2", "K7", "KM"]
            cert_codes = [c for c in optional_codes if c in ["E1", "I1", "IA", "N1", "K8", "E4", "E5", "I5", "K5", "E6", "I6", "C6", "K6", "E7", "I7", "N7", "K7", "IG", "E2", "I2", "IB", "K2", "E3", "I3", "EM", "IM", "KM", "KB", "KD", "KL", "KS", "EP", "IP", "KP"]]
            invalid_certs = [c for c in cert_codes if c not in valid_br6_certs]
            if invalid_certs:
                errors.append(f"BR6 cold temp not compatible with certifications: {', '.join(invalid_certs)}")
        
        # Check for conflicting display options
        display_options = [c for c in optional_codes if c in ["M6", "M5", "M4"]]
        if len(display_options) > 1:
            errors.append(f"Cannot select multiple display options: {', '.join(display_options)}")
        
        # Check for conflicting alarm options
        alarm_options = [c for c in optional_codes if c in ["C4", "CN", "CT", "CR", "CS"]]
        if len(alarm_options) > 1:
            errors.append(f"Cannot select multiple alarm configurations: {', '.join(alarm_options)}")
        
        # Check for conflicting cold temp options
        if "BR5" in optional_codes and "BR6" in optional_codes:
            errors.append("Cannot select both BR5 and BR6 cold temperature options")
        
        return errors
    
    def copy_to_clipboard(self):
        """Copy the generated model number to clipboard"""
        content = self.result_text.get(1.0, tk.END).strip()
        if content and "Complete Model Number:" in content:
            lines = content.split("\n")
            model_num = lines[2] if len(lines) > 2 else ""
            if model_num:
                self.root.clipboard_clear()
                self.root.clipboard_append(model_num)
                messagebox.showinfo("Success", "Model number copied to clipboard!")
        else:
            messagebox.showwarning("No Model Number", "Please generate a model number first!")
    
    def clear_all(self):
        """Clear all selections"""
        for key, data in self.selections.items():
            data["var"].set("")
            data["widget"].set("")
        self.result_text.delete(1.0, tk.END)
        messagebox.showinfo("Cleared", "All selections have been cleared!")

if __name__ == "__main__":
    root = tk.Tk()
    app = Rosemount3051TConfigurator(root)
    root.mainloop()