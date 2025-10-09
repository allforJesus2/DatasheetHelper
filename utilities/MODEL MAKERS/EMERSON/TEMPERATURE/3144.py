import tkinter as tk
from tkinter import ttk

class ScrollableFrame(ttk.Frame):
    """A scrollable frame class for Tkinter."""
    def __init__(self, container, *args, **kwargs):
        super().__init__(container, *args, **kwargs)
        canvas = tk.Canvas(self)
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        self.scrollable_frame = ttk.Frame(canvas)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(
                scrollregion=canvas.bbox("all")
            )
        )

        canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Bind mouse wheel to scroll the canvas
        self.scrollable_frame.bind_all("<MouseWheel>", self._on_mousewheel, add="+")
        canvas.bind_all("<MouseWheel>", self._on_mousewheel, add="+")

    def _on_mousewheel(self, event):
        # This function is bound to the MouseWheel event.
        # It scrolls the canvas content.
        # The direction of scroll is determined by event.delta.
        # This works on Windows and MacOS. For Linux, you might need to bind to <Button-4> and <Button-5>.
        # We find the canvas to scroll.
        canvas = self.winfo_children()[0]
        canvas.yview_scroll(int(-1*(event.delta/120)), "units")

class ModelGeneratorApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Rosemount 3144P Model Number Generator")
        self.geometry("800x600")

        # --- Data extracted from the PDF ---
        self.model_options = {
            "Model": {"3144P": "Temperature transmitter"},
            "Housing Style": {
                "D1": "Field mount, dual-compartment, Aluminum, 1/2-14 NPT",
                "D2": "Field mount, dual-compartment, Aluminum, M20 x 1.5",
                "D5": "Field mount, dual-compartment, Stainless steel, 1/2-14 NPT",
                "D6": "Field mount, dual-compartment, Stainless steel, M20 x 1.5",
            },
            "Transmitter Output": {
                "A": "4-20 mA with HART Protocol",
                "F": "FOUNDATION Fieldbus",
            },
            "Measurement Configuration": {
                "1": "Single-sensor input",
                "2": "Dual-sensor input",
            },
            "Product Certification": {
                "NA": "No approval",
                "E5": "USA explosion-proof, dust ignition-proof, non-incendive",
                "KB": "USA and Canada IS, explosion-proof, non-incendive combination",
                "E1": "ATEX flameproof approval",
                "N1": "ATEX type n approval",
                "K1": "ATEX IS, flameproof, dust ignition-proof, type n combination",
                "ND": "ATEX dust ignition-proof approval",
                "E7": "IECEx flameproof approval",
                "N7": "IECEx Type 'n' approval",
                "I7": "IECEx intrinsic safety",
            },
            "Plantweb/Diagnostics": {
                "": "None",
                "A01": "FOUNDATION Fieldbus advanced control function block suite",
                "D01": "FOUNDATION Fieldbus sensor and process diagnostic suite",
                "DA1": "HART sensor and process diagnostic suite",
            },
            "Enhanced Performance": {
                "": "None",
                "PT": "Temperature measurement assembly with Rosemount X-well Technology",
                "P8": "Enhanced transmitter accuracy",
            },
            "Mounting Bracket": {
                "": "None",
                "B4": "U mounting bracket for 2-in. pipe mounting - all SST",
                "B5": "L mounting bracket for 2-in. pipe or panel mounting - all SST",
                "BH": "L mounting bracket for 2-in. pipe or panel mounting - 316 SST",
            },
            "Display": {"": "None", "M5": "LCD display"},
            "External Ground": {"": "None", "G1": "External ground lug assembly"},
            "Transient Protector": {"": "None", "T1": "Integral transient protector"},
            "Software Configuration": {"": "None", "C1": "Custom configuration"},
            "Line Filter": {"": "None", "F5": "50 Hz line voltage filter"},
            "Alarm Level": {
                "": "None",
                "A1": "NAMUR alarm and saturation levels, high alarm",
                "CN": "NAMUR alarm and saturation levels, low alarm",
            },
            "Low Alarm": {"": "None", "C8": "Low alarm"},
            "Sensor Trim": {"": "None", "C2": "Transmitter-sensor matching"},
            "Calibration": {
                "": "None",
                "C4": "5-point calibration",
                "Q4": "Calibration certificate (3-point)",
            },
            "Dual-Input Config": {
                "": "None",
                "U1": "Hot Backup",
                "U2": "Average temp with Hot Backup and sensor drift alert - Warning",
                "U3": "Average temp with Hot Backup and sensor drift alert - Alarm",
                "U5": "Differential temperature",
                "U6": "Average temperature",
            },
            "Safety Certification": {
                "": "None",
                "QS": "Prior-use certificate of FMEDA data (HART only)",
                "QT": "Safety-certified to IEC 61508 with certificate of FMEDA data (HART only)",
            },
            "Assemble to Sensor": {"": "None", "XA": "Sensor specified separately and assembled to transmitter"},
             "HART Revision": {
                "": "None",
                "HR7": "Configured for HART Revision 7",
            }
        }
        
        self.option_vars = {}
        self.comboboxes = {}

        # --- Main Layout ---
        main_frame = ttk.Frame(self, padding="10")
        main_frame.pack(fill="both", expand=True)
        
        scrollable_window = ScrollableFrame(main_frame)
        scrollable_window.pack(fill="both", expand=True)

        self.create_widgets(scrollable_window.scrollable_frame)

        # --- Bottom Frame for Results and Buttons ---
        bottom_frame = ttk.Frame(self, padding="10")
        bottom_frame.pack(fill="x")
        
        self.result_label = ttk.Label(bottom_frame, text="Generated Model Number:", font=("Helvetica", 12, "bold"))
        self.result_label.pack(pady=5)
        
        self.result_var = tk.StringVar()
        self.result_entry = ttk.Entry(bottom_frame, textvariable=self.result_var, state="readonly", font=("Helvetica", 12))
        self.result_entry.pack(fill="x", pady=5)
        
        button_frame = ttk.Frame(bottom_frame)
        button_frame.pack(pady=10)

        generate_button = ttk.Button(button_frame, text="Generate Model Number", command=self.generate_model_number)
        generate_button.pack(side="left", padx=5)

        clear_button = ttk.Button(button_frame, text="Clear Selections", command=self.clear_selections)
        clear_button.pack(side="left", padx=5)

    def create_widgets(self, parent_frame):
        """Creates and places the dropdown widgets."""
        for i, (category, options) in enumerate(self.model_options.items()):
            label = ttk.Label(parent_frame, text=category, font=("Helvetica", 10, "bold"))
            label.grid(row=i, column=0, sticky="w", padx=10, pady=5)
            
            var = tk.StringVar()
            self.option_vars[category] = var
            
            # Format options as "Code - Description" for display
            formatted_options = []
            for code, description in options.items():
                if code:  # Only add code if it's not empty
                    formatted_options.append(f"{code} - {description}")
                else:
                    formatted_options.append(description)  # For empty codes, just show description
            
            combobox = ttk.Combobox(parent_frame, textvariable=var, values=formatted_options, state="readonly", width=60)
            combobox.grid(row=i, column=1, sticky="ew", padx=10, pady=5)
            combobox.set(formatted_options[0]) # Set default value
            
            self.comboboxes[category] = combobox

            # Special handling for Transmitter Output to control dependent options
            if category == "Transmitter Output":
                var.trace_add("write", self.update_dependent_options)

    def update_dependent_options(self, *args):
        """Enable/disable options based on Transmitter Output selection."""
        selected_display = self.option_vars["Transmitter Output"].get()
        # Extract code from the formatted display (format: "Code - Description")
        selected_code = ""
        if " - " in selected_display:
            selected_code = selected_display.split(" - ")[0]
        else:
            # Fallback for old format or empty codes
            for code, desc in self.model_options["Transmitter Output"].items():
                if desc == selected_display:
                    selected_code = code
                    break
        
        is_hart = selected_code == 'A'
        is_ff = selected_code == 'F'

        # Options unavailable with FOUNDATION Fieldbus
        ff_unavailable = ["Software Configuration", "Alarm Level", "HART Revision", "Safety Certification"]
        # Options unavailable with HART
        hart_unavailable = ["Plantweb/Diagnostics"] # Simplified for this example

        for category, combobox in self.comboboxes.items():
            if category in ff_unavailable:
                combobox.config(state="readonly" if is_hart else "disabled")
                if is_ff:
                    # Reset to first option (None)
                    first_option = combobox['values'][0]
                    self.option_vars[category].set(first_option)
            
            # Example for HART-specific disable
            if category == "Plantweb/Diagnostics":
                 if "DA1" in self.option_vars[category].get():
                      pass # Don't disable if a HART option is selected
                 else:
                    combobox.config(state="readonly" if is_ff else "disabled")
                    if is_hart:
                       # Reset to first option (None)
                       first_option = combobox['values'][0]
                       self.option_vars[category].set(first_option)


    def generate_model_number(self):
        """Assembles the model number from selected options."""
        model_parts = []
        
        # Required components
        required_order = ["Model", "Housing Style", "Transmitter Output", "Measurement Configuration", "Product Certification"]
        
        for category in required_order:
            selected_display = self.option_vars[category].get()
            # Extract code from formatted display
            if " - " in selected_display:
                code = selected_display.split(" - ")[0]
                model_parts.append(code)
            else:
                # Fallback for old format
                for code, desc in self.model_options[category].items():
                    if desc == selected_display:
                        model_parts.append(code)
                        break
        
        # Additional options
        for category, var in self.option_vars.items():
            if category not in required_order:
                selected_display = var.get()
                if selected_display != "None" and not selected_display.startswith("None"):
                    # Extract code from formatted display
                    if " - " in selected_display:
                        code = selected_display.split(" - ")[0]
                        if code:  # Ensure code is not empty
                            model_parts.append(code)
                    else:
                        # Fallback for old format
                        for code, desc in self.model_options[category].items():
                            if desc == selected_display and code:
                                model_parts.append(code)
                                break
                            
        self.result_var.set(" ".join(model_parts))

    def clear_selections(self):
        """Resets all dropdowns to their default value."""
        for category, var in self.option_vars.items():
            # Get the first formatted option (which should be the default)
            first_option = self.comboboxes[category]['values'][0]
            var.set(first_option)
        self.result_var.set("")
        self.update_dependent_options() # Re-check dependencies

if __name__ == "__main__":
    app = ModelGeneratorApp()
    app.mainloop()
