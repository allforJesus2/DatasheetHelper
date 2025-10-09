import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
import json

class Rosemount214CGenerator:
    def __init__(self, root):
        self.root = root
        self.root.title("Rosemount 214C Model Number Generator")
        self.root.geometry("900x700")
        
        # Create main frame with scrollbar
        main_frame = tk.Frame(root)
        main_frame.pack(fill=tk.BOTH, expand=1)
        
        # Create canvas
        canvas = tk.Canvas(main_frame)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=1)
        
        # Add scrollbar
        scrollbar = ttk.Scrollbar(main_frame, orient=tk.VERTICAL, command=canvas.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        
        # Create frame inside canvas
        self.frame = tk.Frame(canvas)
        canvas.create_window((0, 0), window=self.frame, anchor="nw")
        
        # Bind mousewheel to scrollbar
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        
        # Initialize variables
        self.sensor_category = tk.StringVar(value="RTD")
        self.model_parts = {}
        self.optional_parts = {}
        
        # Define configuration data
        self.init_config_data()
        
        # Create UI
        self.create_ui()
        
        # Update initial state
        self.on_category_change()
        
    def init_config_data(self):
        """Initialize configuration data for RTDs and Thermocouples"""
        self.rtd_config = {
            "sensor_type": {
                "RT": "RTD, PT100 (Thin-film, -58 to 842°F)",
                "RW": "RTD, PT100 (Wire-wound, -321 to 1112°F)",
                "RH": "RTD, PT100 (High temp thin-film, -76 to 1112°F)"
            },
            "sheath_material": {
                "SM": "321 SST (Max 1500°F)"
            },
            "accuracy": {
                "A1": "Class A accuracy",
                "B1": "Class B accuracy"
            },
            "elements": {
                "S3": "Single, 3-wire",
                "S4": "Single, 4-wire",
                "D3": "Dual, 3-wire"
            }
        }
        
        self.tc_config = {
            "sensor_type": {
                "TE": "Type E (-40 to 1500°F)",
                "TJ": "Type J (-40 to 1400°F)",
                "TK": "Type K (-40 to 2192°F)",
                "TN": "Type N (-40 to 2192°F)",
                "TT": "Type T (-321 to 698°F)"
            },
            "sheath_material": {
                "SM": "321 SST (for E, J, T)",
                "AK": "Alloy 600 (for K, N)"
            },
            "accuracy": {
                "T1": "Class 1 per IEC 60584",
                "T2": "Class 2 per IEC 60584",
                "SP": "Special Tolerances per ASTM E230",
                "ST": "Standard Tolerances per ASTM E230"
            },
            "elements": {
                "SG": "Single, grounded",
                "SU": "Single, ungrounded",
                "DG": "Dual, grounded, unisolated",
                "DU": "Dual, ungrounded, isolated"
            }
        }
        
        self.common_config = {
            "dimension_units": {
                "E": "English/U.S. (inches)",
                "M": "Metric (mm)"
            },
            "mounting_style": {
                "SL": "Spring-loaded adapter",
                "SC": "Compact spring-loaded adapter",
                "SW": "Spring-loaded with contact indication",
                "WA": "Welded adapter",
                "WC": "Compact welded adapter",
                "SA": "Adjustable spring-loaded fitting",
                "CA": "Compression fitting 1/8-in NPT",
                "CB": "Compression fitting 1/4-in NPT",
                "CC": "Compression fitting 1/2-in NPT",
                "CD": "Compression fitting 3/4-in NPT",
                "DF": "DIN mounting plate with flying leads",
                "DT": "DIN mounting plate with terminal block",
                "SO": "Sensor only"
            }
        }
        
        self.optional_config = {
            "material_options": {
                "M1": "316 SST wire tag",
                "M2": "316 SST components"
            },
            "certification": {
                "E5": "USA Explosionproof",
                "N5": "USA Division 2",
                "E6": "Canada Explosionproof",
                "N6": "Canada Division 2",
                "E1": "ATEX Flameproof",
                "I1": "ATEX Intrinsic Safety",
                "N1": "ATEX Zone 2",
                "E7": "IECEx Flameproof",
                "I7": "IECEx Intrinsic Safety",
                "KB": "USA & Canada Explosionproof"
            },
            "connection_head": {
                "AR1": "Rosemount aluminum",
                "AR2": "Rosemount aluminum with display",
                "SR1": "Rosemount SST",
                "SR2": "Rosemount SST with display",
                "AD1": "Dual entry aluminum",
                "SD1": "Dual entry SST",
                "AF1": "BUZ aluminum",
                "AF3": "BUZH aluminum",
                "AT1": "Aluminum with terminal strip",
                "AJ1": "Universal 3 entry aluminum junction box"
            },
            "conduit_entry": {
                "C1": "1/2-in NPT",
                "C2": "M20 x 1.5",
                "C3": "3/4-in NPT"
            },
            "instrument_connection": {
                "B1": "1/2-in NPT",
                "B2": "M20 x 1.5",
                "B4": "M24 x 1.5"
            },
            "extension_type": {
                "UA": "Union style, 1/2-in NPT",
                "FA": "Fixed style, 1/2-in NPT",
                "PD": "DIN-style, 12x1.5, M24x1.5, M18x1.5",
                "TC": "DIN-style, 12x1.5, M24x1.5, 1/2-in NPT"
            },
            "accessories": {
                "G1": "External ground screw",
                "G3": "Cover chain",
                "TB": "Terminal block",
                "LT": "Low temp connection head (-60°F)",
                "WR3": "3-year warranty",
                "WR5": "5-year warranty"
            }
        }
        
    def create_ui(self):
        """Create the user interface"""
        # Title
        title_label = tk.Label(self.frame, text="Rosemount 214C Temperature Sensor Configurator", 
                               font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=10)
        
        # Sensor Category Selection
        category_frame = tk.LabelFrame(self.frame, text="Sensor Category", font=("Arial", 10, "bold"))
        category_frame.grid(row=1, column=0, columnspan=3, padx=10, pady=5, sticky="ew")
        
        tk.Radiobutton(category_frame, text="RTD", variable=self.sensor_category, 
                      value="RTD", command=self.on_category_change).pack(side=tk.LEFT, padx=10)
        tk.Radiobutton(category_frame, text="Thermocouple", variable=self.sensor_category, 
                      value="TC", command=self.on_category_change).pack(side=tk.LEFT, padx=10)
        
        # Required Components
        req_frame = tk.LabelFrame(self.frame, text="Required Components", font=("Arial", 10, "bold"))
        req_frame.grid(row=2, column=0, columnspan=3, padx=10, pady=5, sticky="ew")
        
        row_num = 0
        
        # Sensor Type
        tk.Label(req_frame, text="Sensor Type:").grid(row=row_num, column=0, sticky="w", padx=5, pady=2)
        self.sensor_type_var = tk.StringVar()
        self.sensor_type_combo = ttk.Combobox(req_frame, textvariable=self.sensor_type_var, width=50)
        self.sensor_type_combo.grid(row=row_num, column=1, padx=5, pady=2)
        self.sensor_type_combo.bind("<<ComboboxSelected>>", self.on_sensor_type_change)
        row_num += 1
        
        # Sheath Material
        tk.Label(req_frame, text="Sheath Material:").grid(row=row_num, column=0, sticky="w", padx=5, pady=2)
        self.sheath_var = tk.StringVar()
        self.sheath_combo = ttk.Combobox(req_frame, textvariable=self.sheath_var, width=50)
        self.sheath_combo.grid(row=row_num, column=1, padx=5, pady=2)
        row_num += 1
        
        # Accuracy
        tk.Label(req_frame, text="Accuracy:").grid(row=row_num, column=0, sticky="w", padx=5, pady=2)
        self.accuracy_var = tk.StringVar()
        self.accuracy_combo = ttk.Combobox(req_frame, textvariable=self.accuracy_var, width=50)
        self.accuracy_combo.grid(row=row_num, column=1, padx=5, pady=2)
        row_num += 1
        
        # Number of Elements
        tk.Label(req_frame, text="Elements:").grid(row=row_num, column=0, sticky="w", padx=5, pady=2)
        self.elements_var = tk.StringVar()
        self.elements_combo = ttk.Combobox(req_frame, textvariable=self.elements_var, width=50)
        self.elements_combo.grid(row=row_num, column=1, padx=5, pady=2)
        row_num += 1
        
        # Dimension Units
        tk.Label(req_frame, text="Dimension Units:").grid(row=row_num, column=0, sticky="w", padx=5, pady=2)
        self.units_var = tk.StringVar()
        self.units_combo = ttk.Combobox(req_frame, textvariable=self.units_var, width=50)
        self.units_combo['values'] = [f"{code} - {desc}" for code, desc in self.common_config['dimension_units'].items()]
        self.units_combo.grid(row=row_num, column=1, padx=5, pady=2)
        self.units_combo.bind("<<ComboboxSelected>>", self.on_units_change)
        row_num += 1
        
        # Insertion Length
        tk.Label(req_frame, text="Insertion Length:").grid(row=row_num, column=0, sticky="w", padx=5, pady=2)
        length_frame = tk.Frame(req_frame)
        length_frame.grid(row=row_num, column=1, sticky="w", padx=5, pady=2)
        self.length_var = tk.StringVar()
        self.length_entry = tk.Entry(length_frame, textvariable=self.length_var, width=10)
        self.length_entry.pack(side=tk.LEFT)
        self.length_label = tk.Label(length_frame, text="inches")
        self.length_label.pack(side=tk.LEFT, padx=5)
        row_num += 1
        
        # Mounting Style
        tk.Label(req_frame, text="Mounting Style:").grid(row=row_num, column=0, sticky="w", padx=5, pady=2)
        self.mounting_var = tk.StringVar()
        self.mounting_combo = ttk.Combobox(req_frame, textvariable=self.mounting_var, width=50)
        self.mounting_combo['values'] = [f"{code} - {desc}" for code, desc in self.common_config['mounting_style'].items()]
        self.mounting_combo.grid(row=row_num, column=1, padx=5, pady=2)
        
        # Optional Components
        opt_frame = tk.LabelFrame(self.frame, text="Optional Components", font=("Arial", 10, "bold"))
        opt_frame.grid(row=3, column=0, columnspan=3, padx=10, pady=5, sticky="ew")
        
        row_num = 0
        
        # Material Options
        tk.Label(opt_frame, text="Material Options:").grid(row=row_num, column=0, sticky="w", padx=5, pady=2)
        self.material_var = tk.StringVar()
        self.material_combo = ttk.Combobox(opt_frame, textvariable=self.material_var, width=50)
        self.material_combo['values'] = ['None'] + [f"{code} - {desc}" for code, desc in self.optional_config['material_options'].items()]
        self.material_combo.set('None')
        self.material_combo.grid(row=row_num, column=1, padx=5, pady=2)
        row_num += 1
        
        # Certification
        tk.Label(opt_frame, text="Certification:").grid(row=row_num, column=0, sticky="w", padx=5, pady=2)
        self.cert_var = tk.StringVar()
        self.cert_combo = ttk.Combobox(opt_frame, textvariable=self.cert_var, width=50)
        self.cert_combo['values'] = ['None'] + [f"{code} - {desc}" for code, desc in self.optional_config['certification'].items()]
        self.cert_combo.set('None')
        self.cert_combo.grid(row=row_num, column=1, padx=5, pady=2)
        row_num += 1
        
        # Connection Head
        tk.Label(opt_frame, text="Connection Head:").grid(row=row_num, column=0, sticky="w", padx=5, pady=2)
        self.head_var = tk.StringVar()
        self.head_combo = ttk.Combobox(opt_frame, textvariable=self.head_var, width=50)
        self.head_combo['values'] = ['None'] + [f"{code} - {desc}" for code, desc in self.optional_config['connection_head'].items()]
        self.head_combo.set('None')
        self.head_combo.grid(row=row_num, column=1, padx=5, pady=2)
        row_num += 1
        
        # Conduit Entry
        tk.Label(opt_frame, text="Conduit Entry:").grid(row=row_num, column=0, sticky="w", padx=5, pady=2)
        self.conduit_var = tk.StringVar()
        self.conduit_combo = ttk.Combobox(opt_frame, textvariable=self.conduit_var, width=50)
        self.conduit_combo['values'] = ['None'] + [f"{code} - {desc}" for code, desc in self.optional_config['conduit_entry'].items()]
        self.conduit_combo.set('None')
        self.conduit_combo.grid(row=row_num, column=1, padx=5, pady=2)
        row_num += 1
        
        # Extension Type
        tk.Label(opt_frame, text="Extension Type:").grid(row=row_num, column=0, sticky="w", padx=5, pady=2)
        self.ext_type_var = tk.StringVar()
        self.ext_type_combo = ttk.Combobox(opt_frame, textvariable=self.ext_type_var, width=50)
        self.ext_type_combo['values'] = ['None'] + [f"{code} - {desc}" for code, desc in self.optional_config['extension_type'].items()]
        self.ext_type_combo.set('None')
        self.ext_type_combo.grid(row=row_num, column=1, padx=5, pady=2)
        self.ext_type_combo.bind("<<ComboboxSelected>>", self.on_extension_change)
        row_num += 1
        
        # Extension Length
        tk.Label(opt_frame, text="Extension Length:").grid(row=row_num, column=0, sticky="w", padx=5, pady=2)
        ext_frame = tk.Frame(opt_frame)
        ext_frame.grid(row=row_num, column=1, sticky="w", padx=5, pady=2)
        self.ext_length_var = tk.StringVar()
        self.ext_length_entry = tk.Entry(ext_frame, textvariable=self.ext_length_var, width=10, state='disabled')
        self.ext_length_entry.pack(side=tk.LEFT)
        self.ext_length_label = tk.Label(ext_frame, text="inches")
        self.ext_length_label.pack(side=tk.LEFT, padx=5)
        row_num += 1
        
        # Accessories
        tk.Label(opt_frame, text="Accessories:").grid(row=row_num, column=0, sticky="w", padx=5, pady=2)
        self.accessory_var = tk.StringVar()
        self.accessory_combo = ttk.Combobox(opt_frame, textvariable=self.accessory_var, width=50)
        self.accessory_combo['values'] = ['None'] + [f"{code} - {desc}" for code, desc in self.optional_config['accessories'].items()]
        self.accessory_combo.set('None')
        self.accessory_combo.grid(row=row_num, column=1, padx=5, pady=2)
        
        # Generate Button
        generate_btn = tk.Button(self.frame, text="Generate Model Number", command=self.generate_model_number,
                                bg="green", fg="white", font=("Arial", 12, "bold"))
        generate_btn.grid(row=4, column=0, columnspan=3, pady=10)
        
        # Results Frame
        results_frame = tk.LabelFrame(self.frame, text="Generated Model Number", font=("Arial", 10, "bold"))
        results_frame.grid(row=5, column=0, columnspan=3, padx=10, pady=5, sticky="ew")
        
        self.result_text = scrolledtext.ScrolledText(results_frame, height=8, width=80)
        self.result_text.pack(padx=5, pady=5)
        
        # Copy Button
        copy_btn = tk.Button(self.frame, text="Copy Model Number", command=self.copy_model_number)
        copy_btn.grid(row=6, column=0, columnspan=3, pady=5)
        
    def on_category_change(self):
        """Handle sensor category change"""
        if self.sensor_category.get() == "RTD":
            self.sensor_type_combo['values'] = [f"{code} - {desc}" for code, desc in self.rtd_config['sensor_type'].items()]
            self.sheath_combo['values'] = [f"{code} - {desc}" for code, desc in self.rtd_config['sheath_material'].items()]
            self.accuracy_combo['values'] = [f"{code} - {desc}" for code, desc in self.rtd_config['accuracy'].items()]
            self.elements_combo['values'] = [f"{code} - {desc}" for code, desc in self.rtd_config['elements'].items()]
        else:  # Thermocouple
            self.sensor_type_combo['values'] = [f"{code} - {desc}" for code, desc in self.tc_config['sensor_type'].items()]
            self.accuracy_combo['values'] = [f"{code} - {desc}" for code, desc in self.tc_config['accuracy'].items()]
            self.elements_combo['values'] = [f"{code} - {desc}" for code, desc in self.tc_config['elements'].items()]
            # Set initial sheath material options
            self.on_sensor_type_change(None)
            
    def on_sensor_type_change(self, event):
        """Handle sensor type change for thermocouples"""
        if self.sensor_category.get() == "TC":
            sensor_type = self.sensor_type_var.get()
            if "Type K" in sensor_type or "Type N" in sensor_type:
                self.sheath_combo['values'] = ["AK - Alloy 600 (for K, N)"]
                self.sheath_var.set("AK - Alloy 600 (for K, N)")
            else:
                self.sheath_combo['values'] = ["SM - 321 SST (for E, J, T)"]
                self.sheath_var.set("SM - 321 SST (for E, J, T)")
                
    def on_units_change(self, event):
        """Handle dimension units change"""
        if "English" in self.units_var.get():
            self.length_label.config(text="inches (0.25 increments)")
            self.ext_length_label.config(text="inches (0.5 increments)")
        else:
            self.length_label.config(text="mm (1 mm increments)")
            self.ext_length_label.config(text="mm (5 mm increments)")
            
    def on_extension_change(self, event):
        """Handle extension type change"""
        if self.ext_type_var.get() == 'None':
            self.ext_length_entry.config(state='disabled')
            self.ext_length_var.set('')
        else:
            self.ext_length_entry.config(state='normal')
            
    def get_code_from_value(self, value, config_dict):
        """Get code from description value"""
        # Handle the new format with "CODE - Description"
        if " - " in value:
            return value.split(" - ")[0]
        
        # Fallback to old format
        for code, desc in config_dict.items():
            if desc == value:
                return code
        return None
        
    def format_length(self, length_str, unit_type, is_extension=False):
        """Format length according to specifications"""
        try:
            if not length_str:
                return None
                
            length = float(length_str)
            
            if "English" in unit_type:
                # For insertion length: 0.25 inch increments
                # For extension: 0.5 inch increments
                if is_extension:
                    # Round to nearest 0.5
                    length = round(length * 2) / 2
                    if length < 2.5 or length > 20:
                        raise ValueError("Extension length must be between 2.5 and 20 inches")
                else:
                    # Round to nearest 0.25
                    length = round(length * 4) / 4
                    if length < 0 or length > 78.5:
                        raise ValueError("Insertion length must be between 0 and 78.5 inches")
                
                # Format as 4-digit code (drop second decimal)
                return f"{int(length * 10):04d}"
            else:  # Metric
                length = int(length)
                if is_extension:
                    # Round to nearest 5mm
                    length = round(length / 5) * 5
                    if length < 65 or length > 500:
                        raise ValueError("Extension length must be between 65 and 500 mm")
                else:
                    if length < 0 or length > 2000:
                        raise ValueError("Insertion length must be between 0 and 2000 mm")
                
                return f"{length:04d}"
        except (ValueError, TypeError) as e:
            return None
            
    def generate_model_number(self):
        """Generate the complete model number"""
        model_parts = ["214C"]
        descriptions = ["Base Model: 214C Temperature Sensor"]
        errors = []
        
        # Get sensor category config
        if self.sensor_category.get() == "RTD":
            config = self.rtd_config
        else:
            config = self.tc_config
            
        # Sensor Type
        sensor_code = self.get_code_from_value(self.sensor_type_var.get(), config['sensor_type'])
        if sensor_code:
            model_parts.append(sensor_code)
            descriptions.append(f"Sensor Type: {sensor_code} - {self.sensor_type_var.get()}")
        else:
            errors.append("Please select a sensor type")
            
        # Sheath Material
        if self.sensor_category.get() == "RTD":
            sheath_code = "SM"  # Only option for RTD
        else:
            sheath_code = self.get_code_from_value(self.sheath_var.get(), self.tc_config['sheath_material'])
            if not sheath_code:
                if "Alloy" in self.sheath_var.get():
                    sheath_code = "AK"
                else:
                    sheath_code = "SM"
        model_parts.append(sheath_code)
        descriptions.append(f"Sheath Material: {sheath_code} - {self.sheath_var.get()}")
        
        # Accuracy
        accuracy_code = self.get_code_from_value(self.accuracy_var.get(), config['accuracy'])
        if accuracy_code:
            model_parts.append(accuracy_code)
            descriptions.append(f"Accuracy: {accuracy_code} - {self.accuracy_var.get()}")
        else:
            errors.append("Please select accuracy")
            
        # Elements
        elements_code = self.get_code_from_value(self.elements_var.get(), config['elements'])
        if elements_code:
            model_parts.append(elements_code)
            descriptions.append(f"Elements: {elements_code} - {self.elements_var.get()}")
        else:
            errors.append("Please select number of elements")
            
        # Dimension Units
        units_code = self.get_code_from_value(self.units_var.get(), self.common_config['dimension_units'])
        if units_code:
            model_parts.append(units_code)
            descriptions.append(f"Units: {units_code} - {self.units_var.get()}")
        else:
            errors.append("Please select dimension units")
            
        # Insertion Length
        length_code = self.format_length(self.length_var.get(), self.units_var.get())
        if length_code:
            model_parts.append(length_code)
            descriptions.append(f"Insertion Length: {length_code} - {self.length_var.get()} {self.length_label['text']}")
        else:
            errors.append("Please enter a valid insertion length")
            
        # Mounting Style
        mounting_code = self.get_code_from_value(self.mounting_var.get(), self.common_config['mounting_style'])
        if mounting_code:
            model_parts.append(mounting_code)
            descriptions.append(f"Mounting: {mounting_code} - {self.mounting_var.get()}")
        else:
            errors.append("Please select a mounting style")
            
        # Optional components
        optional_parts = []
        
        # Material Options
        if self.material_var.get() != 'None':
            mat_code = self.get_code_from_value(self.material_var.get(), self.optional_config['material_options'])
            if mat_code:
                optional_parts.append(mat_code)
                descriptions.append(f"Material Option: {mat_code} - {self.material_var.get()}")
                
        # Certification
        if self.cert_var.get() != 'None':
            cert_code = self.get_code_from_value(self.cert_var.get(), self.optional_config['certification'])
            if cert_code:
                optional_parts.append(cert_code)
                descriptions.append(f"Certification: {cert_code} - {self.cert_var.get()}")
                
        # Connection Head
        if self.head_var.get() != 'None':
            head_code = self.get_code_from_value(self.head_var.get(), self.optional_config['connection_head'])
            if head_code:
                optional_parts.append(head_code)
                descriptions.append(f"Connection Head: {head_code} - {self.head_var.get()}")
                
        # Conduit Entry
        if self.conduit_var.get() != 'None':
            conduit_code = self.get_code_from_value(self.conduit_var.get(), self.optional_config['conduit_entry'])
            if conduit_code:
                optional_parts.append(conduit_code)
                descriptions.append(f"Conduit Entry: {conduit_code} - {self.conduit_var.get()}")
                
        # Extension Type and Length
        if self.ext_type_var.get() != 'None':
            ext_code = self.get_code_from_value(self.ext_type_var.get(), self.optional_config['extension_type'])
            if ext_code:
                optional_parts.append(ext_code)
                descriptions.append(f"Extension Type: {ext_code} - {self.ext_type_var.get()}")
                
                # Extension Length
                if self.ext_length_var.get():
                    ext_length_code = self.format_length(self.ext_length_var.get(), self.units_var.get(), is_extension=True)
                    if ext_length_code:
                        optional_parts.append(f"E{ext_length_code[1:]}")  # Remove leading 0, add E prefix
                        descriptions.append(f"Extension Length: E{ext_length_code[1:]} - {self.ext_length_var.get()} {self.ext_length_label['text']}")
                        
        # Accessories
        if self.accessory_var.get() != 'None':
            acc_code = self.get_code_from_value(self.accessory_var.get(), self.optional_config['accessories'])
            if acc_code:
                optional_parts.append(acc_code)
                descriptions.append(f"Accessory: {acc_code} - {self.accessory_var.get()}")
                
        # Display results
        if errors:
            messagebox.showerror("Configuration Error", "\n".join(errors))
            return
            
        # Build final model number
        model_number = " ".join(model_parts)
        if optional_parts:
            model_number += " " + " ".join(optional_parts)
            
        # Display in result text
        self.result_text.delete(1.0, tk.END)
        self.result_text.insert(tk.END, f"Generated Model Number: {model_number}\n\n")
        self.result_text.insert(tk.END, "Configuration Details:\n")
        for desc in descriptions:
            self.result_text.insert(tk.END, f"• {desc}\n")
            
    def copy_model_number(self):
        """Copy the generated model number to clipboard"""
        content = self.result_text.get(1.0, tk.END).strip()
        if content and "Generated Model Number:" in content:
            # Extract just the model number
            lines = content.split("\n")
            model_num = ""
            for line in lines:
                if line.startswith("Generated Model Number:"):
                    model_num = line.split(": ", 1)[1]
                    break
            
            if model_num:
                self.root.clipboard_clear()
                self.root.clipboard_append(model_num)
                messagebox.showinfo("Success", "Model number copied to clipboard!")
        else:
            messagebox.showwarning("No Model Number", "Please generate a model number first!")

if __name__ == "__main__":
    root = tk.Tk()
    app = Rosemount214CGenerator(root)
    root.mainloop()