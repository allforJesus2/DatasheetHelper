import tkinter as tk
from tkinter import ttk

# Data extracted from the Rosemount Manifold Solutions PDF.
# This structure holds all the options for building a model number.
MANIFOLD_DATA = {
    "R306 (In-line Manifold)": {
        "prefix": "R306",
        "required": [
            {"name": "Design Class", "options": {"Enhanced": "E"}},
            {"name": "Manifold Style", "options": {"Threaded": "T"}},
            {"name": "Manifold Type", "options": {"Block-and-Bleed": "1", "Two-valve": "2"}},
            {"name": "Body Material", "options": {"316 SST": "2"}, "description": "Body, Bonnet, Stem, and Drain/Vent are all 316 SST."},
            {
                "name": "Process Connection",
                "options": {
                    "1/2\"-14 male NPT process x 1/2\"-14 male NPT instrument": "AA",
                    "1/2\"-14 male ANPT process x 1/2\"-14 female ANPT instrument": "AW",
                    "1/2\"-14 female NPT process x 1/2\"-14 male NPT instrument": "BA",
                    "1/2\"-14 female ANPT process x 1/2\"-14 female ANPT instrument": "BW",
                    "3/4\"-14 male NPT process x 1/2\"-14 male NPT instrument": "CA",
                }
            },
            {"name": "Packing Material", "options": {"PTFE": "1", "Graphite-based": "2"}, "description": "PTFE option includes PTFE tape on drain/vent valves and plugs.\nGraphite option includes graphite tape."},
            {"name": "Valve Seat", "options": {"Integral": "1"}},
        ],
        "additional": [
            {"name": "Warranty", "options": {"None": "", "3-year limited": "WR3", "5-year limited": "WR5"}},
            {"name": "Swivel Gauge Adapter", "options": {"Not Included": "", "Included": "SW"}},
            {"name": "Cleaning", "options": {"Standard": "", "Cleaning for special services": "P2"}, "description": "Not available with graphite-based packing material."},
            {"name": "NACE Material Rec.", "options": {"None": "", "Sour Gas (NACE MR0175/ISO 15156, MR0103)": "SG", "Sour Gas w/ 316 SST non-wetted": "TI"}, "description": "Stem and tip material is C-276."},
            {"name": "Material Traceability", "options": {"None": "", "EN 10204 3.1": "Q8"}},
            {"name": "NACE Certificate", "options": {"None": "", "Compliance to NACE MR0175/ISO 15156": "Q15", "Compliance to NACE MR0103": "Q25"}},
        ]
    },
    "0306 (Pressure Manifold)": {
        "prefix": "0306",
        "required": [
            {"name": "Manufacturer", "options": {"Rosemount": "R"}},
            {"name": "Manifold Style", "options": {"Threaded": "T"}},
            {"name": "Manifold Type", "options": {"Block-and-bleed": "1", "Two-valve": "2", "Two-valve (ASME B31.1/B31.3)": "3"}},
            {
                "name": "Material of Construction",
                "options": {
                    "316/316L SST": "2",
                    "Alloy C-276": "3",
                    "Alloy 400": "4",
                    "Alloy 625": "8",
                    "All super duplex SST (UNS S32760)": "9",
                },
                "description": "See datasheet for details on NACE compliance per material."
            },
            {
                "name": "Process Connection",
                "options": {
                    "1/2\"-14 male NPT process x 1/2\"-14 male NPT instrument": "AA",
                    "1/2\"-14 male ANPT process x 1/2\"-14 female ANPT instrument": "AW",
                    "1/2\"-14 female NPT process x 1/2\"-14 male NPT instrument": "BA",
                    "1/2\"-14 female ANPT process x 1/2\"-14 female ANPT instrument": "BW",
                }
            },
            {"name": "Packing Material", "options": {"PTFE": "1", "Graphite-based": "2"}, "description": "PTFE option includes PTFE tape on drain/vent valves and plugs.\nGraphite option includes graphite tape on plugs."},
            {"name": "Valve Seat", "options": {"Integral": "1"}},
        ],
        "additional": [
            {"name": "Warranty", "options": {"None": "", "3-year limited": "WR3", "5-year limited": "WR5"}},
            {"name": "Cleaning", "options": {"Standard": "", "Cleaning for special services": "P2"}, "description": "Not available with graphite-based packing material."},
            {"name": "Cold Temperature", "options": {"None": "", "-67°F (-55°C)": "CW1", "-76°F (-60°C)": "BR6"}, "description": "Only available with specific configurations. See datasheet."},
            {"name": "NACE Material Rec.", "options": {"None": "", "Sour gas (meets NACE MR0175/ISO 15156, MR0103)": "SG", "Sour Gas w/ 316 SST non-wetted": "TI"}, "description": "Only allowed with 316/316L SST material."},
            {"name": "Material Traceability", "options": {"None": "", "EN 10204 3.1": "Q8"}},
            {"name": "NACE Certificate", "options": {"None": "", "Compliance to NACE MR0175/ISO 15156": "Q15", "Compliance to NACE MR0103": "Q25"}},
             {"name": "Swivel Gauge Adapter", "options": {"Not Included": "", "Included": "SW"}},
        ]
    }
}


class ManifoldBuilderApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Rosemount Manifold Model Builder")
        self.geometry("800x700")
        self.configure(bg="#f0f0f0")

        self.option_widgets = []
        self.current_model_data = None
        self.use_hyphens = tk.BooleanVar(value=False)
        
        # --- Top Frame for Model Selection and Display ---
        top_frame = ttk.Frame(self, padding="10")
        top_frame.pack(fill=tk.X, side=tk.TOP)

        ttk.Label(top_frame, text="Select Manifold Model:", font=("Arial", 10, "bold")).pack(side=tk.LEFT, padx=(0, 10))
        
        self.model_selector = ttk.Combobox(top_frame, values=list(MANIFOLD_DATA.keys()), state="readonly", width=30)
        self.model_selector.pack(side=tk.LEFT)
        self.model_selector.bind("<<ComboboxSelected>>", self.populate_options)

        # --- Frame for the generated model number ---
        model_display_frame = ttk.Frame(self, padding="10 5")
        model_display_frame.pack(fill=tk.X, side=tk.TOP)

        ttk.Label(model_display_frame, text="Generated Model Number:", font=("Arial", 10, "bold")).pack(side=tk.LEFT)
        self.model_number_var = tk.StringVar()
        model_number_label = ttk.Label(model_display_frame, textvariable=self.model_number_var, font=("Courier", 12, "bold"), foreground="blue")
        model_number_label.pack(side=tk.LEFT, padx=10)
        
        # Copy button
        copy_button = ttk.Button(model_display_frame, text="Copy", command=self.copy_to_clipboard, width=8)
        copy_button.pack(side=tk.LEFT, padx=5)
        
        # Toggle hyphens button
        self.hyphen_button = ttk.Button(model_display_frame, text="Add Hyphens", command=self.toggle_hyphens, width=12)
        self.hyphen_button.pack(side=tk.LEFT, padx=5)

        # --- Scrollable Area for Options ---
        self.create_scrollable_area()
        
        # Select the first model by default to populate the UI
        self.model_selector.current(0)
        self.populate_options()

    def create_scrollable_area(self):
        """Creates the canvas and scrollbar for the dynamic options."""
        container = ttk.Frame(self, borderwidth=1, relief="sunken")
        container.pack(fill=tk.BOTH, expand=True, side=tk.TOP, padx=10, pady=(0, 10))
        
        self.canvas = tk.Canvas(container, bg="#ffffff")
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=self.canvas.yview)
        
        self.scrollable_frame = ttk.Frame(self.canvas, padding="15")
        
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(
                scrollregion=self.canvas.bbox("all")
            )
        )

        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=scrollbar.set)

        self.canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Bind mouse wheel scrolling
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)

    def _on_mousewheel(self, event):
        """Scrolls the canvas with the mouse wheel."""
        # On Windows, event.delta is typically +/- 120
        self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")

    def populate_options(self, event=None):
        """Clears and rebuilds the UI with dropdowns for the selected model."""
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
        
        self.option_widgets = []
        selected_model = self.model_selector.get()
        self.current_model_data = MANIFOLD_DATA.get(selected_model)

        if not self.current_model_data:
            return

        # --- Required Components ---
        req_label = ttk.Label(self.scrollable_frame, text="Required Model Components", font=("Arial", 14, "bold", "underline"))
        req_label.grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 10))
        
        row_counter = 1
        for option_info in self.current_model_data["required"]:
            has_description = self.create_option_widget(option_info, row_counter)
            row_counter += 2 if has_description else 1

        # --- Additional Options ---
        ttk.Separator(self.scrollable_frame, orient='horizontal').grid(row=row_counter, column=0, columnspan=2, sticky='ew', pady=20)
        row_counter += 1

        add_label = ttk.Label(self.scrollable_frame, text="Additional Options", font=("Arial", 14, "bold", "underline"))
        add_label.grid(row=row_counter, column=0, columnspan=2, sticky="w", pady=(0, 10))
        row_counter += 1

        for option_info in self.current_model_data["additional"]:
            has_description = self.create_option_widget(option_info, row_counter)
            row_counter += 2 if has_description else 1
            
        self.update_model_number()
        # Update canvas scroll region after populating
        self.scrollable_frame.update_idletasks()
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))


    def create_option_widget(self, option_info, row):
        """Creates a label, combobox, and description for a single model option."""
        name = option_info["name"]
        options = option_info["options"]
        description = option_info.get("description", "")
        
        # Label for the option category (e.g., "Manifold Style")
        label = ttk.Label(self.scrollable_frame, text=f"{name}:", font=("Arial", 10, "bold"))
        label.grid(row=row, column=0, sticky="w", padx=(0, 10), pady=(5,0))

        # Format combo values to show "Code - Description"
        combo_values = [f"{code} - {desc}" for desc, code in options.items()]
        
        # Dropdown with the choices
        combo = ttk.Combobox(self.scrollable_frame, values=combo_values, state="readonly", width=50)
        combo.grid(row=row, column=1, sticky="ew", pady=(5,0))
        combo.current(0) # Select the first item by default
        combo.bind("<<ComboboxSelected>>", self.update_model_number)
        
        self.option_widgets.append((combo, options))
        
        # Description label if one exists
        if description:
            desc_label = ttk.Label(self.scrollable_frame, text=description, wraplength=450, justify=tk.LEFT, foreground="gray")
            desc_label.grid(row=row + 1, column=1, sticky="w", pady=(0, 10), padx=5)
            # This is a bit of a trick to increment the main row counter from inside
            # We return True if a description was added so the main loop can skip a row
            return True 
        return False

    def update_model_number(self, event=None):
        """Gathers selections and constructs the full model number string."""
        if not self.current_model_data:
            self.model_number_var.set("")
            return

        required_codes = []
        for combo, options_map in self.option_widgets[:len(self.current_model_data["required"])]:
            selection = combo.get()
            # Extract code from "Code - Description" format
            code = selection.split(" - ")[0] if " - " in selection else selection
            required_codes.append(code)
            
        additional_codes = []
        for combo, options_map in self.option_widgets[len(self.current_model_data["required"]):]:
            selection = combo.get()
            # Extract code from "Code - Description" format
            code = selection.split(" - ")[0] if " - " in selection else selection
            if code: # Only add non-empty codes for additional options
                additional_codes.append(code)

        # Build model number with or without hyphens
        if self.use_hyphens.get():
            model_number = self.current_model_data["prefix"] + "-" + "-".join(required_codes)
            if additional_codes:
                model_number += " " + "-".join(additional_codes)
        else:
            model_number = self.current_model_data["prefix"] + "".join(required_codes)
            if additional_codes:
                model_number += " " + "".join(additional_codes)
            
        self.model_number_var.set(model_number)
    
    def copy_to_clipboard(self):
        """Copies the model number to the clipboard."""
        model_number = self.model_number_var.get()
        if model_number:
            self.clipboard_clear()
            self.clipboard_append(model_number)
            self.update()  # Required for clipboard to work
            # Visual feedback - change button text briefly
            self.after(100, lambda: self.flash_copy_feedback())
    
    def flash_copy_feedback(self):
        """Provides brief visual feedback that copy was successful."""
        # Could add a status message or tooltip here if desired
        pass
    
    def toggle_hyphens(self):
        """Toggles between hyphenated and non-hyphenated model numbers."""
        current = self.use_hyphens.get()
        self.use_hyphens.set(not current)
        # Update button text
        if self.use_hyphens.get():
            self.hyphen_button.config(text="Remove Hyphens")
        else:
            self.hyphen_button.config(text="Add Hyphens")
        # Regenerate model number
        self.update_model_number()


if __name__ == "__main__":
    app = ManifoldBuilderApp()
    app.mainloop()
