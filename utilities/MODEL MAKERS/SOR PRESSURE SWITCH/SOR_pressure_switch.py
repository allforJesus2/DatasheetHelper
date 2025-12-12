import tkinter as tk
from tkinter import ttk

# Model options compiled from product data (CAT456)
# Descriptions are based on examples like 6AG-EF3-M4-C2A and 5BH-JF45-M1-C2A
MODEL_OPTIONS = {
    "Model": {
        "6": "Model 6 (e.g., Piston/Spring)",
        "5": "Model 5 (e.g., Diaphragm/Spring)",
        # Add more models as needed
    },
    "Housing": {
        "AG": "AG (Aluminum, Hazardous Loc)",
        "AH": "AH (Stainless Steel, Hazardous Loc)",
        "BH": "BH (Stainless Steel, ATEX/IECEx)",
        "BG": "BG (Housing Type BG)",
    },
    "Switching Element": {
        "EF": "EF (SPDT, Normal AC/DC)",
        "JF": "JF (SPDT, Gold Contacts, Low Power)",
        # Add more elements as needed
    },
    "Model Suffix (Range)": {
        "3": "Range 3 (e.g., 12-100 psi w/ Model 6)",
        "5": "Range 5 (e.g., 20-180 psi w/ Model 6)",
        "45": "Range 45 (e.g., 45-550 psi w/ Model 5)",
        # Add more ranges as needed
    },
    "Diaphragm/O-Ring": {
        "M1": "M1 (Viton GLT O-Ring, 316L SS Diaphragm)",
        "M4": "M4 (Viton O-Ring, 316L SS Diaphragm)",
        "N4": "N4 (Buna-N O-Ring, 316L SS Diaphragm)",
    },
    "Pressure Port": {
        "C1A": "C1A (1/4\" NPT-F, 316SS)",
        "C2A": "C2A (1/2\" NPT-F, 316SS)",
    },
    "Accessories": {
        "YY": "YY (No Accessories)",
        "C4": "C4 (Certificate of Compliance)",
        "A1": "A1 (Certificate of Calibration)",
        "HBMETTA1": "HBMETTA1 (Terminal Box + SS Tag)",
        "X": "X (Special Accessory)",
    }
}

class ModelConfigurator(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("SOR Mini-Hermet Configurator (CAT456)")
        self.geometry("750x450")
        
        # Set a style
        self.style = ttk.Style(self)
        self.style.theme_use('clam')
        
        # Dictionary to hold the StringVars
        self.vars = {}
        
        # Dictionary to hold the description labels
        self.desc_labels = {}

        self.create_widgets()
        self.update_model_number() # Initialize the model number display

    def create_widgets(self):
        main_frame = ttk.Frame(self, padding="15")
        main_frame.pack(fill="both", expand=True)

        options_frame = ttk.Frame(main_frame)
        options_frame.pack(fill="x", expand=True)

        row = 0
        for name, choices in MODEL_OPTIONS.items():
            # Label for the component
            ttk.Label(options_frame, text=f"{name}:", font=("Helvetica", 10, "bold")).grid(
                row=row, column=0, sticky="w", padx=5, pady=8
            )
            
            # String Variable to hold the selection
            self.vars[name] = tk.StringVar(self)
            
            # Combobox (dropdown) for options (show as "CODE - description")
            choice_list = [f"{code} - {desc}" for code, desc in choices.items()]
            combo = ttk.Combobox(
                options_frame, 
                textvariable=self.vars[name], 
                values=choice_list, 
                width=15,
                state="readonly"
            )
            combo.grid(row=row, column=1, sticky="ew", padx=5, pady=8)
            
            # Set default selection
            self.vars[name].set(choice_list[0])
            
            # Description Label
            # Extract code part for description lookup
            _first_code = choice_list[0].split(" - ", 1)[0]
            desc_text = choices.get(_first_code, "No description.")
            desc_label = ttk.Label(
                options_frame, 
                text=desc_text, 
                wraplength=400, 
                font=("Helvetica", 9), 
                foreground="#333"
            )
            desc_label.grid(row=row, column=2, sticky="w", padx=10)
            self.desc_labels[name] = desc_label
            
            # Bind the combobox selection event to update description and model number
            combo.bind(
                "<<ComboboxSelected>>",
                lambda event, n=name, ch=choices: self.update_description_and_model(n, ch)
            )
            
            row += 1
            
        options_frame.columnconfigure(2, weight=1) # Allow description to expand

        # --- Result Display ---
        ttk.Separator(main_frame, orient="horizontal").pack(fill="x", pady=20)
        
        result_frame = ttk.Frame(main_frame)
        result_frame.pack(fill="both", expand=True)

        ttk.Label(result_frame, text="Configured Model Number:", font=("Helvetica", 12, "bold")).pack(pady=(0, 5))
        
        self.result_var = tk.StringVar()
        self.result_label = ttk.Label(
            result_frame, 
            textvariable=self.result_var, 
            font=("Courier", 16, "bold"), 
            background="white", 
            padding=10, 
            relief="solid",
            anchor="center"
        )
        self.result_label.pack(fill="x", pady=5)

    def update_description_and_model(self, name, choices_dict, *args):
        """Update the description label and the final model number."""
        selected_display = self.vars[name].get()
        selected_key = selected_display.split(" - ", 1)[0] if " - " in selected_display else selected_display
        
        # Update description
        if name in self.desc_labels:
            self.desc_labels[name].config(text=choices_dict.get(selected_key, "No description"))
            
        # Update model number
        self.update_model_number()

    def update_model_number(self, *args):
        """Concatenate the selected options to form the final model number."""
        try:
            # Extract just the code part before " - " for each selection
            def _code(value: str) -> str:
                return value.split(" - ", 1)[0] if " - " in value else value

            p1 = _code(self.vars["Model"].get())
            p2 = _code(self.vars["Housing"].get())
            p3 = _code(self.vars["Switching Element"].get())
            p4 = _code(self.vars["Model Suffix (Range)"].get())
            p5 = _code(self.vars["Diaphragm/O-Ring"].get())
            p6 = _code(self.vars["Pressure Port"].get())
            p7 = _code(self.vars["Accessories"].get())
            
            # Format: 6AG-EF3-M4-C2A-YY
            # The structure is: [P1][P2]-[P3][P4]-[P5]-[P6]-[P7]
            
            base_model = f"{p1}{p2}-{p3}{p4}-{p5}-{p6}"
            
            # Only add accessories if not 'YY' (None)
            if p7 == "YY":
                final_model = base_model
            else:
                final_model = f"{base_model}-{p7}"
                
            self.result_var.set(final_model)
            
        except KeyError:
            # This can happen during initialization before all vars are created
            self.result_var.set("Initializing...")
        except Exception as e:
            print(f"Error updating model: {e}")
            self.result_var.set("Error")

if __name__ == "__main__":
    app = ModelConfigurator()
    app.mainloop()
