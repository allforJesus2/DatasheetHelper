import tkinter as tk
from tkinter import ttk, messagebox
import re

class ThermowellConfigurator(tk.Tk):
    def __init__(self):
        super().__init__()
        
        self.title("Rosemount 114C Thermowell Configurator")
        self.geometry("900x700")
        
        # Create main container with scrollbar
        main_frame = ttk.Frame(self)
        main_frame.pack(fill=tk.BOTH, expand=1)
        
        # Create canvas
        canvas = tk.Canvas(main_frame)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=1)
        
        # Add scrollbar
        scrollbar = ttk.Scrollbar(main_frame, orient=tk.VERTICAL, command=canvas.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Configure canvas
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        
        # Mouse wheel scrolling
        canvas.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(int(-1*(e.delta/120)), "units"))
        
        # Create frame inside canvas
        self.config_frame = ttk.Frame(canvas)
        canvas.create_window((0, 0), window=self.config_frame, anchor="nw")
        
        # Model number display at top
        display_frame = ttk.LabelFrame(self, text="Generated Model Number", padding="10")
        display_frame.pack(fill=tk.X, padx=10, pady=10)
        
        self.model_var = tk.StringVar(value="114C-")
        model_label = ttk.Label(display_frame, textvariable=self.model_var, 
                               font=("Courier", 14, "bold"), foreground="blue")
        model_label.pack()
        
        # Copy button
        copy_btn = ttk.Button(display_frame, text="Copy Model Number", command=self.copy_model)
        copy_btn.pack(pady=5)
        
        # Store selections
        self.selections = {}
        self.option_vars = {}  # Store option checkbox variables
        
        # Build configuration options
        self.create_widgets()
        
    def create_widgets(self):
        # Model (Fixed)
        self.add_section("Model", 0)
        self.add_fixed_field("Model", "114C")
        
        # Units
        self.add_section("Dimension Units", 1)
        units_options = [
            ("E - English units (in.)", "E"),
            ("M - Metric units (mm)", "M")
        ]
        self.add_dropdown("Units", units_options, self.update_model)
        
        # Immersion Length
        self.add_section("Immersion Length (U)", 2)
        self.add_entry("Immersion", "Enter 4-digit length (e.g., 0100 for 10in or 0250 for 250mm)", 4)
        
        # Mounting Style
        self.add_section("Mounting Style", 3)
        mounting_options = [
            ("T - Threaded", "T"),
            ("P - Flanged, partial penetration weld", "P"),
            ("F - Flanged, full penetration weld", "F"),
            ("G - Flanged, forged (no welds)", "G"),
            ("V - Van Stone, lap flange", "V"),
            ("W - Welded, socket weld", "W"),
            ("D - Welded, weld-in", "D")
        ]
        self.add_dropdown("Mounting", mounting_options, self.update_mounting_options)
        
        # Process Connection
        self.add_section("Process Connection", 4)
        self.process_frame = ttk.Frame(self.config_frame)
        self.process_frame.grid(row=self.current_row, column=0, columnspan=2, sticky="ew", padx=20, pady=5)
        self.process_dropdown = None
        self.current_row += 1
        
        # Stem Style
        self.add_section("Stem Style", 5)
        stem_options = [
            ("1 - Straight", "1"),
            ("2 - Tapered", "2"),
            ("3 - Stepped", "3")
        ]
        self.add_dropdown("Stem", stem_options, self.update_model)
        
        # Thermowell Material
        self.add_section("Thermowell Material", 6)
        material_options = [
            ("SC - 316/316L SST", "SC"),
            ("SF - 304/304L SST", "SF"),
            ("CS - Carbon steel (A-105)", "CS"),
            ("SD - 316/316L SST (NORSOK)", "SD"),
            ("SG - 316Ti SST", "SG"),
            ("AC - Alloy C-276", "AC"),
            ("AN - Alloy 625", "AN"),
            ("DS - Super duplex SST", "DS"),
            ("DU - Duplex 2205", "DU"),
            ("TT - Titanium Grade 2", "TT")
        ]
        self.add_dropdown("Material", material_options, self.update_model)
        
        # Head Length
        self.add_section("Head Length (H)", 7)
        self.add_entry("Head", "Enter 3-digit length (e.g., 045 for 45mm or 017 for 1.75in)", 3)
        
        # Instrument Connection
        self.add_section("Instrument Connection", 8)
        instrument_options = [
            ("A - 1/2-14 NPT", "A"),
            ("B - 1/2-14 NPSM", "B"),
            ("C - 3/4-14 NPT", "C"),
            ("D - M18 x 1.5p", "D"),
            ("E - M20 x 1.5p", "E"),
            ("F - M24 x 1.5p", "F"),
            ("G - G 1/2-in. (BSPF)", "G"),
            ("H - G 3/4-in. (BSPF)", "H"),
            ("J - M27 x 2p", "J"),
            ("K - M14 x 1.5p", "K")
        ]
        self.add_dropdown("Instrument", instrument_options, self.update_model)
        
        # Options - Multiple checkboxes
        self.add_section("Additional Options (Optional)", 9)
        self.add_option_checkboxes()
        
        # Initialize dropdown
        self.update_mounting_options()
        
    def add_section(self, title, row):
        self.current_row = row * 3
        label = ttk.Label(self.config_frame, text=title, font=("Arial", 10, "bold"))
        label.grid(row=self.current_row, column=0, sticky="w", padx=20, pady=(10, 2))
        
    def add_fixed_field(self, key, value):
        self.selections[key] = value
        label = ttk.Label(self.config_frame, text=value, foreground="gray")
        label.grid(row=self.current_row + 1, column=0, sticky="w", padx=40)
        
    def add_dropdown(self, key, options, callback=None):
        var = tk.StringVar()
        self.selections[key] = var
        
        combo = ttk.Combobox(self.config_frame, textvariable=var, width=50, state="readonly")
        combo['values'] = [opt[0] for opt in options]
        combo.grid(row=self.current_row + 1, column=0, columnspan=2, sticky="ew", padx=40, pady=2)
        
        # Store mapping
        setattr(self, f"{key}_map", {opt[0]: opt[1] for opt in options})
        
        if callback:
            combo.bind('<<ComboboxSelected>>', lambda e: callback())
        else:
            combo.bind('<<ComboboxSelected>>', lambda e: self.update_model())
            
        if options:
            combo.current(0)
            callback() if callback else self.update_model()
    
    def add_entry(self, key, placeholder, width, required=True):
        var = tk.StringVar()
        self.selections[key] = var
        
        entry = ttk.Entry(self.config_frame, textvariable=var, width=width)
        entry.grid(row=self.current_row + 1, column=0, sticky="w", padx=40, pady=2)
        entry.insert(0, placeholder)
        entry.config(foreground="gray")
        
        def on_focus_in(e):
            if entry.get() == placeholder:
                entry.delete(0, tk.END)
                entry.config(foreground="black")
        
        def on_focus_out(e):
            if entry.get() == "":
                entry.insert(0, placeholder)
                entry.config(foreground="gray")
            self.update_model()
        
        entry.bind("<FocusIn>", on_focus_in)
        entry.bind("<FocusOut>", on_focus_out)
        entry.bind("<KeyRelease>", lambda e: self.update_model())
    
    def add_option_checkboxes(self):
        # Common options with descriptions
        common_options = [
            ("Q5", "Standard external pressure test"),
            ("Q9", "Extended external pressure test"),
            ("Q8", "Material certification"),
            ("Q35", "NACE approval"),
            ("Q73", "Dye penetration test"),
            ("Q76", "PMI testing"),
            ("Q83", "Ultrasonic wall thickness test"),
            ("Q84", "Radiography (X-ray) wall thickness test"),
            ("Q85", "Standard internal pressure test"),
            ("Q86", "Extended internal pressure test"),
            ("R21", "Thermowell calculation"),
            ("R14", "Surface finish < Ra 0.3 μm"),
            ("R20", "Electropolish"),
            ("R60", "Spherical tip"),
            ("R11", "Vent hole"),
            ("WR3", "3-year limited warranty"),
            ("WR5", "5-year limited warranty"),
            ("XT", "Hand tight assembly"),
            ("XW", "Process-ready assembly")
        ]
        
        # Create a frame for checkboxes with grid layout
        options_frame = ttk.Frame(self.config_frame)
        options_frame.grid(row=self.current_row + 1, column=0, columnspan=2, sticky="ew", padx=40, pady=5)
        
        # Create checkboxes in 2 columns
        for idx, (code, desc) in enumerate(common_options):
            var = tk.BooleanVar()
            self.option_vars[code] = var
            
            cb = ttk.Checkbutton(
                options_frame, 
                text=f"{code} - {desc}",
                variable=var,
                command=self.update_model
            )
            row = idx // 2
            col = idx % 2
            cb.grid(row=row, column=col, sticky="w", padx=10, pady=2)
        
        # Configure column weights for better spacing
        options_frame.columnconfigure(0, weight=1)
        options_frame.columnconfigure(1, weight=1)
    
    def update_mounting_options(self):
        mounting = self.selections.get("Mounting")
        if not mounting:
            return
            
        mounting_val = self.Mounting_map.get(mounting.get(), "T")
        

        # Check if process_frame exists before trying to clear it
        if not hasattr(self, 'process_frame'):
            return
        
        # Clear existing dropdown
        for widget in self.process_frame.winfo_children():
            widget.destroy()
        
        # Define process connections based on mounting type
        if mounting_val == "T":
            options = [
                ("AA - 1/2-14 NPT", "AA"),
                ("AB - 3/4-14 NPT", "AB"),
                ("AC - 1-11.5 NPT", "AC"),
                ("AD - 1 1/2-11.5 NPT", "AD"),
                ("DA - M20 x 1.5p", "DA"),
                ("DB - M24 x 1.5p", "DB"),
                ("DE - 1/2-in. BSPF (G1/2)", "DE")
            ]
        elif mounting_val in ["P", "F", "G"]:
            options = [
                ("AA - 1-in. Class 150", "AA"),
                ("AB - 1 1/2-in. Class 150", "AB"),
                ("AC - 2-in. Class 150", "AC"),
                ("AH - 1-in. Class 300", "AH"),
                ("AJ - 1 1/2-in. Class 300", "AJ"),
                ("AK - 2-in. Class 300", "AK")
            ]
        elif mounting_val == "V":
            options = [
                ("AA - 1-in. Class 150", "AA"),
                ("AB - 1 1/2-in. Class 150", "AB"),
                ("AC - 2-in. Class 150", "AC"),
                ("AH - 1-in. Class 300", "AH"),
                ("AJ - 1 1/2-in. Class 300", "AJ"),
                ("AK - 2-in. Class 300", "AK")
            ]
        elif mounting_val in ["W", "D"]:
            options = [
                ("AA - 3/4-in. pipe", "AA"),
                ("AB - 1-in. pipe", "AB"),
                ("AC - 1 1/4-in pipe", "AC"),
                ("AD - 1 1/2-in. pipe", "AD")
            ]
        else:
            options = [("AA - 1/2-14 NPT", "AA")]
        
        var = tk.StringVar()
        self.selections["Process"] = var
        
        combo = ttk.Combobox(self.process_frame, textvariable=var, width=50, state="readonly")
        combo['values'] = [opt[0] for opt in options]
        combo.pack(fill=tk.X, padx=20)
        
        self.Process_map = {opt[0]: opt[1] for opt in options}
        combo.bind('<<ComboboxSelected>>', lambda e: self.update_model())
        
        if options:
            combo.current(0)
        
        self.process_dropdown = combo
        self.update_model()
    
    def update_model(self):
        try:
            model = "114C"
            
            # Units
            units = self.selections.get("Units")
            if units and units.get():
                model += "-" + self.Units_map.get(units.get(), "E")
            else:
                model += "-E"
            
            # Immersion
            immersion = self.selections.get("Immersion")
            if immersion and immersion.get() and immersion.get()[0].isdigit():
                val = immersion.get()[:4].zfill(4)
                model += "-" + val
            else:
                model += "-XXXX"
            
            # Mounting
            mounting = self.selections.get("Mounting")
            if mounting and mounting.get():
                model += "-" + self.Mounting_map.get(mounting.get(), "T")
            else:
                model += "-T"
            
            # Process
            process = self.selections.get("Process")
            if process and process.get():
                model += "-" + self.Process_map.get(process.get(), "AA")
            else:
                model += "-XX"
            
            # Stem
            stem = self.selections.get("Stem")
            if stem and stem.get():
                model += "-" + self.Stem_map.get(stem.get(), "1")
            else:
                model += "-1"
            
            # Material
            material = self.selections.get("Material")
            if material and material.get():
                model += "-" + self.Material_map.get(material.get(), "SC")
            else:
                model += "-SC"
            
            # Head
            head = self.selections.get("Head")
            if head and head.get() and head.get()[0].isdigit():
                val = head.get()[:3].zfill(3)
                model += "-" + val
            else:
                model += "-XXX"
            
            # Instrument
            instrument = self.selections.get("Instrument")
            if instrument and instrument.get():
                model += "-" + self.Instrument_map.get(instrument.get(), "A")
            else:
                model += "-A"
            
            # Options - gather selected checkboxes
            selected_options = [code for code, var in self.option_vars.items() if var.get()]
            if selected_options:
                model += "-" + "-".join(selected_options)
            
            self.model_var.set(model)
            
        except Exception as e:
            print(f"Error updating model: {e}")
    
    def copy_model(self):
        model = self.model_var.get()
        self.clipboard_clear()
        self.clipboard_append(model)
        messagebox.showinfo("Copied", f"Model number copied to clipboard:\n{model}")

if __name__ == "__main__":
    app = ThermowellConfigurator()
    app.mainloop()