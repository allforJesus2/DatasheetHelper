import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import json

class RosemountConfigurator:
    def __init__(self, root):
        self.root = root
        self.root.title("Rosemount 1199 Seal System Configurator")
        self.root.geometry("900x700")
        
        # Model number components storage
        self.selections = {}
        self.mount_type = tk.StringVar(value="Remote")  # Track mount type
        
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
        title = ttk.Label(scrollable_frame, text="Rosemount 1199 Seal System Configurator", 
                         font=("Arial", 16, "bold"))
        title.grid(row=0, column=0, columnspan=3, pady=10, sticky="w")
        
        # Instructions
        instructions = ttk.Label(scrollable_frame, 
                                text="Select options below to generate a complete model number",
                                font=("Arial", 10))
        instructions.grid(row=1, column=0, columnspan=3, pady=5, sticky="w")
        
        row = 2
        
        # Model (Base)
        row = self.add_dropdown(scrollable_frame, row, "Model", "model",
            {"1199": "Seal systems"},
            "Base model number for seal systems")
        
        # Mount Type Selector
        mount_label = ttk.Label(scrollable_frame, text="Mount Type:", font=("Arial", 10, "bold"))
        mount_label.grid(row=row, column=0, sticky="w", pady=5, padx=(0,10))
        
        mount_frame = ttk.Frame(scrollable_frame)
        mount_frame.grid(row=row, column=1, sticky="w", pady=5)
        
        remote_radio = ttk.Radiobutton(mount_frame, text="Remote Mount", variable=self.mount_type, 
                                      value="Remote", command=self.on_mount_type_change)
        remote_radio.pack(side=tk.LEFT, padx=5)
        
        direct_radio = ttk.Radiobutton(mount_frame, text="Direct Mount", variable=self.mount_type,
                                       value="Direct", command=self.on_mount_type_change)
        direct_radio.pack(side=tk.LEFT, padx=5)
        
        mount_desc = ttk.Label(scrollable_frame, 
                              text="Select Remote Mount (with capillary) or Direct Mount (no capillary)",
                              font=("Arial", 8), foreground="gray")
        mount_desc.grid(row=row+1, column=1, sticky="w", pady=(0,5))
        row += 2
        
        # Connection Type (will be updated based on mount type)
        self.connection_row = row
        row = self.add_connection_type_dropdown(scrollable_frame, row)
        
        # Seal Fill Fluid
        row = self.add_dropdown(scrollable_frame, row, "Seal Fill Fluid", "fill_fluid",
            {
                "D": "Silicone 200 (-49 to 401°F)",
                "F": "Silicone 200 for vacuum",
                "J": "Tri-Therm 300 (-40 to 572°F) [Food grade]",
                "Q": "Tri-Therm 300 for vacuum [Food grade]",
                "L": "Silicone 704 (32 to 599°F)",
                "C": "Silicone 704 for vacuum",
                "R": "Silicone 705 (68 to 698°F)",
                "V": "Silicone 705 for vacuum",
                "A": "SYLTHERM XLT (-157 to 293°F)",
                "H": "Inert (Halocarbon) (-49 to 320°F)",
                "G": "Glycerine and water (5 to 203°F) [Food grade]",
                "N": "Neobee M-20 (5 to 437°F) [Food grade]",
                "P": "Propylene Glycol (5 to 203°F) [Food grade]"
            },
            "Fill fluid type and temperature range")
        
        # Seal Connection Type (Direct Mount only)
        self.seal_conn_row = row
        row = self.add_seal_connection_type_dropdown(scrollable_frame, row)
        
        # Seal Connection Type/Capillary ID or Direct Mount Connection
        self.capillary_row = row
        row = self.add_capillary_or_direct_dropdown(scrollable_frame, row)
        
        # Capillary Length (only for Remote Mount)
        self.cap_length_row = row
        row = self.add_capillary_length_dropdown(scrollable_frame, row)
        
        # Industry Standard
        row = self.add_dropdown(scrollable_frame, row, "Industry Standard", "standard",
            {
                "A": "ASME B16.5 (American)",
                "D": "EN 1092-1 (European)",
                "T": "GOST 33259-15 (Russian)",
                "J": "JIS B2238 (Japanese)",
                "G": "HG20615 (Chinese-ASME based)",
                "K": "HG20592 (Chinese-EN based)",
                "S": "Hygienic (3-A Standard 74-06)",
                "N": "Non-industry standard"
            },
            "Manufacturing standard to follow")
        
        # Seal Assembly Type
        row = self.add_dropdown(scrollable_frame, row, "Seal Assembly Type", "seal_type",
            {
                "FFW": "Flush Flanged Seal",
                "RFW": "Remote Flanged Seal",
                "EFW": "Extended Flanged Seal",
                "PFW": "Pancake Seal",
                "RTW": "Remote Threaded Seal",
                "HTS": "Male Threaded Seal",
                "SCW": "Hygienic Tri-Clover Style",
                "SSW": "Hygienic Tank Spud",
                "VCS": "Tri-Clamp In-Line Seal",
                "TFS": "Wafer Style In-Line Seal",
                "WSP": "Saddle Seal"
            },
            "Type of seal assembly")
        
        # Process Connection Size
        row = self.add_dropdown(scrollable_frame, row, "Process Connection Size", "conn_size",
            {
                "1": "½-in. / DN 15",
                "2": "1-in. / DN 25 / 25A",
                "4": "1½-in. / DN 40 / 40A",
                "G": "2-in. / DN 50 / 50A",
                "7": "3-in. / DN 80 / 80A",
                "9": "4-in. / DN 100 / 100A",
                "A": "¾-in. / DN 10 / 10A",
                "B": "DN 15 / 15A",
                "C": "DN 20",
                "D": "DN 25",
                "F": "DN 40",
                "J": "DN 80"
            },
            "Size of process connection")
        
        # Flange/Pressure Rating
        row = self.add_dropdown(scrollable_frame, row, "Pressure Rating", "pressure",
            {
                "1": "Class 150 / 10K",
                "2": "Class 300 / 20K",
                "4": "Class 600 / 40K",
                "5": "Class 900",
                "6": "Class 1500",
                "7": "Class 2500",
                "G": "PN 40",
                "E": "PN 10/16",
                "H": "PN 63",
                "J": "PN 100",
                "K": "PN 160",
                "0": "Based on customer flange"
            },
            "Flange pressure rating")
        
        # Diaphragm Material
        row = self.add_dropdown(scrollable_frame, row, "Diaphragm Material", "diaphragm",
            {
                "CA": "316L SST / CS flange",
                "DA": "316L SST / 316 SST flange",
                "CB": "Alloy C-276 / CS flange",
                "DB": "Alloy C-276 / 316 SST flange",
                "DC": "Tantalum / 316 SST flange",
                "LA": "316L SST / None",
                "LB": "Alloy C-276 / None",
                "LC": "Tantalum / None",
                "D6": "Duplex 2205 SST",
                "D5": "Duplex 2507 SST",
                "RH": "Titanium Gr. 4",
                "DH": "Titanium Gr. 4 / 316L upper"
            },
            "Diaphragm and flange material")
        
        # Lower Housing Material
        row = self.add_dropdown(scrollable_frame, row, "Lower Housing", "lower_housing",
            {
                "0": "None",
                "A": "316L SST",
                "B": "Alloy C-276",
                "2": "Duplex 2205 SST",
                "H": "Titanium Gr. 4",
                "V": "Alloy 400",
                "6": "Nickel 201",
                "F": "304L SST"
            },
            "Lower housing (flushing ring) material")
        
        # Flushing Connections
        row = self.add_dropdown(scrollable_frame, row, "Flushing Connections", "flushing",
            {
                "0": "None",
                "5": "None",
                "1": "One connection (¼-18 NPT)",
                "3": "Two connections (¼-18 NPT)",
                "7": "One connection (½-14 NPT)",
                "9": "Two connections (½-14 NPT)",
                "Y": "Rosemount 319 Flushing Ring"
            },
            "Number and size of flushing connections")
        
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
        
        # Extended Warranty
        row = self.add_dropdown(scrollable_frame, row, "Extended Warranty (Optional)", "warranty",
            {
                "": "None (Standard warranty)",
                "WR3": "3-year limited warranty",
                "WR5": "5-year limited warranty"
            },
            "Extended product warranty coverage")
        
        # Intermediate Gasket Material
        row = self.add_dropdown(scrollable_frame, row, "Intermediate Gasket (Optional)", "gasket",
            {
                "": "None (Default: Klingersil C-4401)",
                "0": "No gasket for flushing ring",
                "Y": "Klingersil C-4401 gasket",
                "J": "PTFE gasket",
                "N": "GRAFOIL gasket",
                "K": "Barium sulfate filled PTFE gasket",
                "R": "Ethylene propylene gasket"
            },
            "Gasket material for lower housing")
        
        # Flushing Plug/Vent Valve
        row = self.add_dropdown(scrollable_frame, row, "Flushing Plug/Vent (Optional)", "plug",
            {
                "": "None",
                "D": "Alloy C-276 plug(s) for flushing connection(s)",
                "G": "316 SST plug(s) for flushing connection(s)",
                "H": "316 SST vent/drain for flushing connection(s)"
            },
            "Plugs or vent/drain valves for flushing connections")
        
        # Low Side Drain/Vent Valve
        row = self.add_dropdown(scrollable_frame, row, "Low Side Drain/Vent (Optional)", "low_vent",
            {
                "": "None",
                "FV": "Low side drain/vent valve (Required for DP measurement)"
            },
            "Drain/vent valve for low pressure side")
        
        # Diaphragm Thickness
        row = self.add_dropdown(scrollable_frame, row, "Diaphragm Thickness (Optional)", "dia_thick",
            {
                "": "Standard thickness",
                "C": "0.006-in. (150 μm) - For abrasive applications",
                "7": "0.002-in. (50 μm) - Thin diaphragm"
            },
            "Special diaphragm thickness for specific applications")
        
        # Diaphragm Coating
        row = self.add_dropdown(scrollable_frame, row, "Diaphragm Coating (Optional)", "coating",
            {
                "": "None (Bare diaphragm)",
                "Z": "0.0002-in. (5 μm) gold plated diaphragm",
                "V": "PTFE coated diaphragm (nonstick)",
                "FP": "CorrosionShield PFA coated diaphragm"
            },
            "Special coatings for diaphragm protection")
        
        # NACE Certificate
        row = self.add_dropdown(scrollable_frame, row, "NACE Certificate (Optional)", "nace",
            {
                "": "None",
                "Q15": "NACE MR0175/ISO 15156 (Sour oil field)",
                "Q25": "NACE MR0103 (Sour refining)"
            },
            "Certificate of compliance for sour service environments")
        
        # Cold Temperature Application
        row = self.add_dropdown(scrollable_frame, row, "Cold Temperature (Optional)", "cold_temp",
            {
                "": "Standard fill",
                "B": "Extra fill for cold temperature application"
            },
            "Additional fill fluid for cold temperature applications")
        
        # Gasket Surface Finish
        row = self.add_dropdown(scrollable_frame, row, "Gasket Surface Finish (Optional)", "surface",
            {
                "": "Standard finish",
                "1": "Gasket surface Ra 125 Max./EN 1092-1 Type B2",
                "D": "10 μin. (0.25 μm) Ra surface finish",
                "G": "15 μin. (0.375 μm) Ra surface finish",
                "H": "20 μin. (0.50 μm) Ra surface finish"
            },
            "Special surface finish requirements")
        
        # Capillary Weld Protection
        row = self.add_dropdown(scrollable_frame, row, "Capillary Weld Protection (Optional)", "weld_prot",
            {
                "": "None",
                "FB": "Environmental corrosion protection for capillary welds"
            },
            "Additional corrosion protection for capillary welds")
        
        # Bolt Material
        row = self.add_dropdown(scrollable_frame, row, "Bolt Material (Optional)", "bolt",
            {
                "": "Default (Tin plated CS)",
                "3": "304 SST bolts",
                "4": "316 SST bolts",
                "FA": "316 SST bolts (stud bolt design)"
            },
            "Upgraded bolt material")
        
        # PMI Verification
        row = self.add_dropdown(scrollable_frame, row, "PMI Verification (Optional)", "pmi",
            {
                "": "None",
                "Q76": "PMI Verification and Certificate"
            },
            "Positive Material Identification verification")
        
        # Lower Housing Alignment
        row = self.add_dropdown(scrollable_frame, row, "Lower Housing Clamp (Optional)", "align",
            {
                "": "None",
                "SA": "Lower housing alignment clamp"
            },
            "Alignment clamp for lower housing")
        
        # Alternate Design
        row = self.add_dropdown(scrollable_frame, row, "Alternate Design (Optional)", "alt_design",
            {
                "": "Standard design",
                "E": "One-piece design",
                "2": "Radial capillary connection",
                "4": "Flat face, flush flanged",
                "9": "Male threads / Large diaphragm (4.1-in.)"
            },
            "Alternative design configurations")
        
        # Special Threading
        row = self.add_dropdown(scrollable_frame, row, "Special Features (Optional)", "special",
            {
                "": "None",
                "JA": "Threaded jack bolt holes in flange",
                "P": "Non-hygienic fill fluid",
                "6": "Electropolishing"
            },
            "Additional special features or modifications")
        
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
        
        self.result_text = tk.Text(scrollable_frame, height=4, width=80, 
                                   font=("Courier", 11), wrap=tk.WORD)
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
        dropdown = ttk.Combobox(parent, textvariable=var, width=50, state="readonly")
        dropdown['values'] = [f"{code} - {desc}" for code, desc in options.items()]
        dropdown.grid(row=row, column=1, sticky="w", pady=5)
        
        # Store reference
        self.selections[key] = {"var": var, "options": options, "widget": dropdown, "label": label, "desc": None}
        
        # Description
        desc_label = ttk.Label(parent, text=description, font=("Arial", 8), 
                              foreground="gray")
        desc_label.grid(row=row+1, column=1, sticky="w", pady=(0,5))
        self.selections[key]["desc"] = desc_label
        
        return row + 2
    
    def add_connection_type_dropdown(self, parent, row):
        """Add connection type dropdown that changes based on mount type"""
        # Label
        label = ttk.Label(parent, text="Connection Type:", font=("Arial", 10, "bold"))
        label.grid(row=row, column=0, sticky="w", pady=5, padx=(0,10))
        
        # Dropdown
        var = tk.StringVar()
        dropdown = ttk.Combobox(parent, textvariable=var, width=50, state="readonly")
        dropdown.grid(row=row, column=1, sticky="w", pady=5)
        
        # Store reference
        self.selections["connection"] = {"var": var, "options": {}, "widget": dropdown, "label": label, "desc": None}
        
        # Description
        desc_label = ttk.Label(parent, text="", font=("Arial", 8), foreground="gray")
        desc_label.grid(row=row+1, column=1, sticky="w", pady=(0,5))
        self.selections["connection"]["desc"] = desc_label
        
        # Initialize with Remote Mount options
        self.update_connection_type_options()
        
        return row + 2
    
    def add_capillary_or_direct_dropdown(self, parent, row):
        """Add capillary type (Remote) or direct mount connection (Direct) dropdown"""
        # Label
        label = ttk.Label(parent, text="", font=("Arial", 10, "bold"))
        label.grid(row=row, column=0, sticky="w", pady=5, padx=(0,10))
        
        # Dropdown
        var = tk.StringVar()
        dropdown = ttk.Combobox(parent, textvariable=var, width=50, state="readonly")
        dropdown.grid(row=row, column=1, sticky="w", pady=5)
        
        # Store reference
        self.selections["capillary"] = {"var": var, "options": {}, "widget": dropdown, "label": label, "desc": None}
        
        # Description
        desc_label = ttk.Label(parent, text="", font=("Arial", 8), foreground="gray")
        desc_label.grid(row=row+1, column=1, sticky="w", pady=(0,5))
        self.selections["capillary"]["desc"] = desc_label
        
        # Initialize with Remote Mount options
        self.update_capillary_or_direct_options()
        
        return row + 2
    
    def add_seal_connection_type_dropdown(self, parent, row):
        """Add seal connection type dropdown (only for Direct Mount)"""
        # Label
        label = ttk.Label(parent, text="Seal Connection Type:", font=("Arial", 10, "bold"))
        label.grid(row=row, column=0, sticky="w", pady=5, padx=(0,10))
        
        # Dropdown
        var = tk.StringVar()
        dropdown = ttk.Combobox(parent, textvariable=var, width=50, state="readonly")
        seal_conn_options = {
            "A": "Direct mount"
        }
        dropdown['values'] = [f"{code} - {desc}" for code, desc in seal_conn_options.items()]
        dropdown.grid(row=row, column=1, sticky="w", pady=5)
        
        # Store reference
        self.selections["seal_conn"] = {"var": var, "options": seal_conn_options, "widget": dropdown, 
                                        "label": label, "desc": None}
        
        # Description
        desc_label = ttk.Label(parent, text="Seal connection type (Direct Mount only)",
                              font=("Arial", 8), foreground="gray")
        desc_label.grid(row=row+1, column=1, sticky="w", pady=(0,5))
        self.selections["seal_conn"]["desc"] = desc_label
        
        # Initially hidden (only shown for Direct Mount)
        label.grid_remove()
        dropdown.grid_remove()
        desc_label.grid_remove()
        
        return row + 2
    
    def add_capillary_length_dropdown(self, parent, row):
        """Add capillary length dropdown (only for Remote Mount)"""
        # Label
        label = ttk.Label(parent, text="Capillary Length:", font=("Arial", 10, "bold"))
        label.grid(row=row, column=0, sticky="w", pady=5, padx=(0,10))
        
        # Dropdown
        var = tk.StringVar()
        dropdown = ttk.Combobox(parent, textvariable=var, width=50, state="readonly")
        cap_length_options = {
            "01": "1.0 ft. (0.3 m)",
            "05": "5.0 ft. (1.5 m)",
            "10": "10.0 ft. (3.0 m)",
            "15": "15.0 ft. (4.5 m)",
            "20": "20.0 ft. (6.1 m)",
            "25": "25.0 ft. (7.6 m)",
            "30": "30.0 ft. (9.1 m)",
            "40": "40.0 ft. (12.2 m)",
            "50": "50.0 ft. (15.2 m)",
            "51": "1.6 ft. (0.5 m)",
            "52": "3.3 ft. (1.0 m)",
            "53": "4.9 ft. (1.5 m)",
            "54": "6.6 ft. (2.0 m)",
            "55": "8.2 ft. (2.5 m)",
            "56": "9.8 ft. (3.0 m)",
            "57": "11.5 ft. (3.5 m)",
            "58": "13.1 ft. (4.0 m)",
            "59": "16.4 ft. (5.0 m)",
            "60": "19.7 ft. (6.0 m)"
        }
        dropdown['values'] = [f"{code} - {desc}" for code, desc in cap_length_options.items()]
        dropdown.grid(row=row, column=1, sticky="w", pady=5)
        
        # Store reference
        self.selections["cap_length"] = {"var": var, "options": cap_length_options, "widget": dropdown, 
                                         "label": label, "desc": None}
        
        # Description
        desc_label = ttk.Label(parent, text="Length of capillary tubing (Remote Mount only)",
                              font=("Arial", 8), foreground="gray")
        desc_label.grid(row=row+1, column=1, sticky="w", pady=(0,5))
        self.selections["cap_length"]["desc"] = desc_label
        
        return row + 2
    
    def update_connection_type_options(self):
        """Update connection type options based on mount type"""
        mount_type = self.mount_type.get()
        
        if mount_type == "Remote":
            options = {
                "W": "Welded-repairable (High side)",
                "M": "Welded-repairable (Low side)",
                "D": "Welded-repairable (Balanced system)",
                "A": "All welded, capillary (High side)",
                "B": "All welded, capillary (Two seal system)",
                "C": "All welded, capillary (Low side)"
            }
            description = "Type of seal system and location on transmitter"
        else:  # Direct Mount
            options = {
                "W": "Welded-repairable (High side) - Coplanar devices",
                "R": "All welded, one seal (High side) - Coplanar devices",
                "T": "All welded, two seal (High side) - Coplanar devices",
                "W_INLINE": "All welded, one seal - In-line devices"
            }
            description = "Connection type and seal location (varies by transmitter type)"
        
        self.selections["connection"]["options"] = options
        self.selections["connection"]["widget"]['values'] = [f"{code} - {desc}" for code, desc in options.items()]
        self.selections["connection"]["desc"].config(text=description)
        self.selections["connection"]["var"].set("")
    
    def update_capillary_or_direct_options(self):
        """Update capillary/direct mount connection options based on mount type"""
        mount_type = self.mount_type.get()
        
        if mount_type == "Remote":
            options = {
                "B": "0.03-in. (0.711 mm) ID",
                "C": "0.04-in. (1.092 mm) ID",
                "D": "0.075-in. (1.905 mm) ID",
                "E": "0.03-in. ID, PVC coated with closed end",
                "F": "0.04-in. ID, PVC coated with closed end",
                "G": "0.075-in. ID, PVC coated with closed end",
                "H": "0.03-in. ID, 4-in. support tube",
                "J": "0.04-in. ID, 4-in. support tube",
                "K": "0.075-in. ID, 4-in. support tube",
                "M": "0.03-in. ID, PVC coated, 4-in. support tube",
                "N": "0.04-in. ID, PVC coated, 4-in. support tube",
                "P": "0.075-in. ID, PVC coated, 4-in. support tube"
            }
            label_text = "Capillary Type"
            description = "Capillary internal diameter and coating"
        else:  # Direct Mount
            options = {
                # Coplanar one-seal system (Welded-repairable)
                "93": "Direct mount, no extension - Welded-repairable, Coplanar one-seal",
                "B3": "Direct mount, 2-in. (50 mm) extension - Welded-repairable, Coplanar one-seal",
                "D3": "Direct mount, 4-in. (100 mm) extension - Welded-repairable, Coplanar one-seal",
                # Coplanar one-seal system (All welded)
                "97": "Direct mount, no extension - All welded, Coplanar one-seal",
                "B7": "Direct mount, 2-in. (50 mm) extension - All welded, Coplanar one-seal",
                "D7": "Direct mount, 4-in. (100 mm) extension - All welded, Coplanar one-seal",
                # Coplanar Tuned-System (Welded-repairable)
                "94": "Direct mount, no extension - Welded-repairable, Tuned-System",
                "B4": "Direct mount, 2-in. (50 mm) extension - Welded-repairable, Tuned-System",
                "D4": "Direct mount, 4-in. (100 mm) extension - Welded-repairable, Tuned-System",
                # Coplanar Tuned-System (All welded)
                "96": "Direct mount, no extension - All welded, Tuned-System",
                "B6": "Direct mount, 2-in. (50 mm) extension - All welded, Tuned-System",
                "D6": "Direct mount, 4-in. (100 mm) extension - All welded, Tuned-System",
                # In-line devices
                "95": "Direct mount, no extension - All welded, In-line one-seal",
                "C5": "Direct mount, 4-in. (100 mm) extension - All welded, In-line one-seal",
                "D5": "Direct mount, Thermal Optimizer - All welded, In-line one-seal"
            }
            label_text = "Direct Mount Connection Type"
            description = "Direct mount connection type with extension length (includes seal system type)"
        
        self.selections["capillary"]["options"] = options
        self.selections["capillary"]["widget"]['values'] = [f"{code} - {desc}" for code, desc in options.items()]
        self.selections["capillary"]["label"].config(text=f"{label_text}:")
        self.selections["capillary"]["desc"].config(text=description)
        self.selections["capillary"]["var"].set("")
    
    def on_mount_type_change(self):
        """Handle mount type change - update relevant dropdowns"""
        mount_type = self.mount_type.get()
        
        # Update connection type options
        self.update_connection_type_options()
        
        # Update capillary/direct mount options
        self.update_capillary_or_direct_options()
        
        # Show/hide fields based on mount type
        if mount_type == "Remote":
            # Hide seal connection type (Direct Mount only)
            self.selections["seal_conn"]["label"].grid_remove()
            self.selections["seal_conn"]["widget"].grid_remove()
            self.selections["seal_conn"]["desc"].grid_remove()
            self.selections["seal_conn"]["var"].set("")
            
            # Show capillary length
            self.selections["cap_length"]["label"].grid()
            self.selections["cap_length"]["widget"].grid()
            self.selections["cap_length"]["desc"].grid()
        else:  # Direct Mount
            # Show seal connection type
            self.selections["seal_conn"]["label"].grid()
            self.selections["seal_conn"]["widget"].grid()
            self.selections["seal_conn"]["desc"].grid()
            # Auto-select "A" if not already set
            if not self.selections["seal_conn"]["var"].get():
                self.selections["seal_conn"]["var"].set("A - Direct mount")
            
            # Hide capillary length
            self.selections["cap_length"]["label"].grid_remove()
            self.selections["cap_length"]["widget"].grid_remove()
            self.selections["cap_length"]["desc"].grid_remove()
            self.selections["cap_length"]["var"].set("")
    
    def generate_model(self):
        """Generate the complete model number"""
        model_parts = []
        missing = []
        mount_type = self.mount_type.get()
        
        # Required fields in order (varies by mount type)
        if mount_type == "Remote":
            required_order = [
                "model", "connection", "fill_fluid", "capillary", "cap_length",
                "standard", "seal_type", "conn_size", "pressure", "diaphragm",
                "lower_housing", "flushing"
            ]
        else:  # Direct Mount
            required_order = [
                "model", "connection", "fill_fluid", "seal_conn", "capillary",  # seal_conn is "A", capillary is direct mount connection code
                "standard", "seal_type", "conn_size", "pressure", "diaphragm",
                "lower_housing", "flushing"
            ]
        
        # Optional fields in order
        optional_order = [
            "warranty", "gasket", "plug", "low_vent", "dia_thick", "coating",
            "nace", "cold_temp", "surface", "weld_prot", "bolt", "pmi",
            "align", "alt_design", "special"
        ]
        
        # Process required fields
        for key in required_order:
            # Skip cap_length for Direct Mount (it's included in capillary code)
            if mount_type == "Direct" and key == "cap_length":
                continue
                
            value = self.selections[key]["var"].get()
            if value:
                # Extract code (first part before " - ")
                code = value.split(" - ")[0]
                # Handle special case for Direct Mount connection type
                if mount_type == "Direct" and key == "connection" and code == "W_INLINE":
                    code = "W"  # Use W for in-line devices
                model_parts.append(code)
            else:
                # Use field name for missing items
                field_names = {
                    "model": "Model",
                    "connection": "Connection Type",
                    "fill_fluid": "Fill Fluid",
                    "seal_conn": "Seal Connection Type",
                    "capillary": "Capillary Type" if mount_type == "Remote" else "Direct Mount Connection",
                    "cap_length": "Capillary Length",
                    "seal_type": "Seal Assembly Type",
                    "standard": "Industry Standard",
                    "conn_size": "Connection Size",
                    "pressure": "Pressure Rating",
                    "diaphragm": "Diaphragm Material",
                    "lower_housing": "Lower Housing",
                    "flushing": "Flushing Connections"
                }
                missing.append(field_names.get(key, key))
        
        if missing:
            messagebox.showwarning("Incomplete Selection", 
                                  f"Please select values for:\n" + "\n".join(f"• {m}" for m in missing))
            return
        
        # Validate combinations before proceeding
        validation_errors = self.validate_configuration(model_parts)
        if validation_errors:
            messagebox.showerror("Invalid Configuration", 
                               "The following issues were found:\n\n" + 
                               "\n\n".join(f"• {err}" for err in validation_errors))
            return
        
        # Process optional fields - only add if selected
        optional_codes = []
        for key in optional_order:
            value = self.selections[key]["var"].get()
            if value and value != "":
                # Extract code (first part before " - ")
                code = value.split(" - ")[0]
                if code and code != "":  # Only add non-empty codes
                    optional_codes.append(code)
        
        # Validate optional codes
        optional_errors = self.validate_optional_codes(model_parts, optional_codes)
        if optional_errors:
            messagebox.showerror("Invalid Optional Configuration", 
                               "The following issues were found:\n\n" + 
                               "\n\n".join(f"• {err}" for err in optional_errors))
            return
        
        # Join required parts with hyphens
        model_number = "-".join(model_parts)
        
        # Add optional codes if any exist
        if optional_codes:
            model_number += "-" + "-".join(optional_codes)
        
        # Display result
        self.result_text.delete(1.0, tk.END)
        mount_type_display = "Remote Mount" if mount_type == "Remote" else "Direct Mount"
        self.result_text.insert(1.0, f"Complete Model Number ({mount_type_display}):\n\n{model_number}\n\n")
        
        # Add summary
        summary = f"Mount Type: {mount_type_display}\n"
        summary += f"Base Configuration: {len(model_parts)} codes\n"
        if optional_codes:
            summary += f"Optional Codes: {len(optional_codes)} selected\n"
            summary += f"Options: {', '.join(optional_codes)}\n\n"
        else:
            summary += "Optional Codes: None selected\n\n"
        
        self.result_text.insert(tk.END, summary)
        if mount_type == "Direct":
            self.result_text.insert(tk.END, "Note: Direct Mount requires specification of a Rosemount pressure device.\n")
            self.result_text.insert(tk.END, "Add seal system ordering code (B11/B12/S1/S2) to transmitter model.\n")
            self.result_text.insert(tk.END, "See Table 1 in datasheet for correct code per transmitter model.")
        else:
            self.result_text.insert(tk.END, "Note: Additional transmitter model required for complete order.")
    
    def validate_configuration(self, parts):
        """Validate the model configuration for compatibility issues"""
        errors = []
        mount_type = self.mount_type.get()
        
        # Extract key components (adjust indices based on required_order)
        connection_type = parts[1] if len(parts) > 1 else ""
        fill_fluid = parts[2] if len(parts) > 2 else ""
        
        # For Direct Mount, seal_conn is at index 3, capillary at index 4
        # For Remote Mount, capillary is at index 3
        if mount_type == "Direct":
            capillary = parts[4] if len(parts) > 4 else ""
            diaphragm = parts[9] if len(parts) > 9 else ""
        else:  # Remote Mount
            capillary = parts[3] if len(parts) > 3 else ""
            diaphragm = parts[9] if len(parts) > 9 else ""
        
        # Rule 1: All welded connection types require 316L SST or Alloy C-276 diaphragm
        if mount_type == "Remote":
            if connection_type in ["A", "B", "C"]:
                if not any(diaphragm.startswith(x) for x in ["CA", "DA", "CB", "DB", "LA", "LB"]):
                    errors.append("All welded connection types (A, B, C) require 316L SST or Alloy C-276 diaphragm material")
        else:  # Direct Mount
            if connection_type in ["R", "T"]:
                if not any(diaphragm.startswith(x) for x in ["CA", "DA", "CB", "DB", "LA", "LB"]):
                    errors.append("All welded Direct Mount connection types (R, T) require 316L SST or Alloy C-276 diaphragm material")
        
        # Rules 2-4: Only apply to Remote Mount (capillary-related)
        if mount_type == "Remote":
            # Rule 2: PVC coated capillaries cannot exceed 212°F (100°C)
            if capillary in ["E", "F", "G", "M", "N", "P"]:
                high_temp_fluids = ["J", "Q", "L", "C", "R", "V"]  # Fluids that can exceed 212°F
                if fill_fluid in high_temp_fluids:
                    errors.append("PVC coated capillaries (E, F, G, M, N, P) cannot be used with fill fluids exceeding 212°F. Choose non-PVC capillary.")
            
            # Rule 3: Silicone 704 only available with certain capillary types
            if fill_fluid in ["L", "C"]:
                if capillary not in ["C", "D", "F", "G", "J", "K", "N", "P"]:
                    errors.append("Silicone 704 fill fluid (L, C) only available with capillary codes C, D, F, G, J, K, N, P")
            
            # Rule 4: Silicone 705 only available with certain capillary types
            if fill_fluid in ["R", "V"]:
                if capillary not in ["D", "G", "K", "P"]:
                    errors.append("Silicone 705 fill fluid (R, V) only available with capillary codes D, G, K, P")
        
        # Rule 5: Glycerine and Propylene Glycol not suitable for vacuum
        if fill_fluid in ["G", "P"]:
            errors.append("WARNING: Glycerine (G) and Propylene Glycol (P) are not suitable for vacuum applications")
        
        # Rule 6: Titanium RH diaphragm restrictions
        if diaphragm == "RH":
            if mount_type == "Remote":
                if connection_type in ["A", "B", "C"]:
                    errors.append("Titanium Gr. 4 (RH) diaphragm not available with all welded connection types (A, B, C)")
            else:  # Direct Mount
                if connection_type in ["R", "T"]:
                    errors.append("Titanium Gr. 4 (RH) diaphragm not available with all welded Direct Mount connection types (R, T)")
        
        # Rule 7: Direct Mount specific - connection type must match direct mount code
        if mount_type == "Direct":
            # Check if connection type matches the direct mount code
            # W is used for: 93/B3/D3 (welded-repairable coplanar), 94/B4/D4 (welded-repairable Tuned-System), 95/C5/D5 (in-line)
            if capillary.startswith("97") or capillary.startswith("B7") or capillary.startswith("D7"):
                if connection_type != "R":
                    errors.append("Direct mount codes 97/B7/D7 require connection type R (All welded, one seal)")
            elif capillary.startswith("96") or capillary.startswith("B6") or capillary.startswith("D6"):
                if connection_type != "T":
                    errors.append("Direct mount codes 96/B6/D6 require connection type T (All welded, two seal)")
            elif capillary.startswith("93") or capillary.startswith("B3") or capillary.startswith("D3") or \
                 capillary.startswith("94") or capillary.startswith("B4") or capillary.startswith("D4") or \
                 capillary.startswith("95") or capillary.startswith("C5") or capillary.startswith("D5"):
                if connection_type not in ["W", "W_INLINE"]:
                    errors.append("Direct mount codes 93/B3/D3/94/B4/D4/95/C5/D5 require connection type W")
        
        return errors
    
    def validate_optional_codes(self, base_parts, optional_codes):
        """Validate optional codes against base configuration"""
        errors = []
        mount_type = self.mount_type.get()
        
        # Extract key components (adjust index for Direct Mount)
        # seal_type is now at index 6 (after standard at index 5)
        if mount_type == "Direct":
            diaphragm = base_parts[9] if len(base_parts) > 9 else ""
            seal_type = base_parts[6] if len(base_parts) > 6 else ""
        else:  # Remote Mount
            diaphragm = base_parts[9] if len(base_parts) > 9 else ""
            seal_type = base_parts[6] if len(base_parts) > 6 else ""
        
        # Rule 1: Diaphragm coating only available on certain materials
        coating_codes = [c for c in optional_codes if c in ["Z", "V", "FP"]]
        if coating_codes:
            valid_for_coating = ["CA", "DA", "LA", "CB", "DB", "LB"]  # 316L SST, Alloy C-276, Alloy 400
            if not any(diaphragm.startswith(x) for x in valid_for_coating):
                errors.append(f"Diaphragm coating ({', '.join(coating_codes)}) only available on 316L SST, Alloy C-276, or Alloy 400 diaphragms")
        
        # Rule 2: CorrosionShield (FP) not compatible with spiral wound gaskets
        if "FP" in optional_codes:
            # Materials that use spiral wound gaskets
            spiral_wound_materials = ["CA", "DA", "LA"]
            if any(diaphragm.startswith(x) for x in spiral_wound_materials):
                errors.append("CorrosionShield (FP) coating is not compatible with spiral wound gaskets (used with CA, DA, LA materials)")
        
        # Rule 3: Abrasive diaphragm thickness only for certain materials
        if "C" in optional_codes:  # 0.006-in. thickness
            valid_for_thick = ["CA", "DA", "LA", "CB", "DB", "LB", "D6", "L6", "C6"]
            if not any(diaphragm.startswith(x) for x in valid_for_thick):
                errors.append("0.006-in. diaphragm thickness (C) only available with 316L SST, Alloy C-276, or Duplex 2205 SST")
        
        # Rule 4: Thin diaphragm only for certain materials
        if "7" in optional_codes:  # 0.002-in. thickness
            valid_for_thin = ["CA", "DA", "LA", "CB", "DB", "LB"]
            if not any(diaphragm.startswith(x) for x in valid_for_thin):
                errors.append("0.002-in. diaphragm thickness (7) only available with 316L SST or Alloy C-276")
        
        # Rule 5: Tantalum brazed not available with certain options
        if diaphragm in ["C3", "D3"]:
            if "C" in optional_codes:
                errors.append("Tantalum brazed diaphragms (C3, D3) not available with thick diaphragm option (C)")
        
        # Rule 6: One-piece design restrictions
        if "E" in optional_codes:
            restricted_materials = ["CA", "CB", "CC", "C3", "MB", "KB"]
            if any(diaphragm.startswith(x) for x in restricted_materials):
                errors.append("One-piece design (E) not available with two-piece only materials (CA, CB, CC, C3, MB, KB)")
        
        # Rule 7: Lower housing alignment clamp only for certain seal types
        if "SA" in optional_codes:
            if seal_type not in ["FFW", "PFW"]:
                errors.append("Lower housing alignment clamp (SA) only available with FFW (Flush Flanged) or PFW (Pancake) seals")
        
        # Rule 8: Check for conflicting diaphragm thickness options
        thickness_options = [c for c in optional_codes if c in ["C", "7"]]
        if len(thickness_options) > 1:
            errors.append(f"Cannot select multiple diaphragm thickness options: {', '.join(thickness_options)}")
        
        # Rule 9: Check for conflicting coating options
        coating_options = [c for c in optional_codes if c in ["Z", "V", "FP"]]
        if len(coating_options) > 1:
            errors.append(f"Cannot select multiple diaphragm coatings: {', '.join(coating_options)}")
        
        # Rule 10: Electropolishing typically for hygienic applications
        if "6" in optional_codes:
            if seal_type not in ["SCW", "SSW", "VCS", "SHP", "SLS", "EES", "SVS", "STW"]:
                errors.append("WARNING: Electropolishing (6) is typically only for hygienic seal types (S-series)")
        
        return errors
        
    def copy_to_clipboard(self):
        """Copy the generated model number to clipboard"""
        content = self.result_text.get(1.0, tk.END).strip()
        if content and "Complete Model Number:" in content:
            # Extract just the model number
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
    app = RosemountConfigurator(root)
    root.mainloop()