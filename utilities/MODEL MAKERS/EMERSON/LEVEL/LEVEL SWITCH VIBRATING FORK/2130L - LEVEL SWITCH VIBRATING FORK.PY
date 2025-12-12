import tkinter as tk
from tkinter import ttk
import tkinter.font as tkFont
import pyperclip

class Rosemount2130Builder:
    def __init__(self, root):
        self.root = root
        self.root.title("Rosemount 2130 Model Number Builder")
        self.root.geometry("800x600")
        
        # Data structure mapping field names to code-description pairs
        self.data = {
            "Model": [("2130", "Enhanced Vibrating Fork Liquid Level Switch")],
            "Output": [
                ("L", "Direct Load Switching (Mains 2-wire) 20 to 264 Vac, 50/60 Hz, 20 to 60 Vdc, Self-checking"),
                ("P", "PNP/PLC Low Voltage (3-wire) 20 to 60 Vdc, Self-checking"),
                ("D", "Relay DPCO (Double Pole Changeover), 20 to 264 Vac, 50/60 Hz, 20 to 60 Vdc, Self-checking"),
                ("N", "NAMUR, 8 Vdc, Self-checking"),
                ("M", "8/16 mA, Self-checking"),
            ],
            "Housing Material": [
                ("A", "Aluminum"),
                ("S", "Stainless Steel"),
            ],
            "Conduit Entry": [
                ("9", "¾-in. NPT"),
                ("2", "M20"),
            ],
            "Operating Temperature": [
                ("M", "Mid-Range: –40 °F (–40 °C) to 356 °F (180 °C)"),
                ("E", "High: –94 °F (–70 °C) to 500 °F (260 °C)"),
            ],
            "Process Connection/Fork Material": [
                ("S", "316/316L Stainless Steel (1.4401/1.4404)"),
                ("F", "ECTFE Copolymer, Coated 316/316L Stainless Steel (1.4401/1.4404)"),
                ("H", "Alloy C (UNS N10002), Alloy C-276 (UNS N10276)"),
            ],
            "Process Connection Size": [
                ("9", "¾-in. / 19 mm"),
                ("1", "1-in. / 25 mm (DN25)"),
                ("2", "2-in. / 50 mm (DN50)"),
                ("5", "1½-in. / 40 mm (DN40)"),
                ("3", "3-in. / 80 mm (DN80)"),
                ("4", "4-in. / 100 mm (DN100)"),
                ("7", "2½-in. / 65 mm (DN65)"),
            ],
            "Process Connection Rating": [
                ("AA", "ASME B16.5 Class 150 flange"),
                ("AB", "ASME B16.5 Class 300 flange"),
                ("AC", "ASME B16.5 Class 600 flange"),
                ("DA", "EN1092-1 PN 10/16 flange"),
                ("DB", "EN1092-1 PN 25/40 flange"),
                ("DC", "EN1092-1 PN 63 flange"),
                ("DD", "EN1092-1 PN 100 flange"),
                ("NN", "For use with non-flange process connection type"),
            ],
            "Process Connection Type": [
                ("R", "Raised Face (RF) Flange"),
                ("B", "BSPT (R) Thread"),
                ("G", "BSPP (G) Thread"),
                ("N", "NPT Thread"),
                ("P", "BSPP (G) O-ring"),
                ("C", "Tri-Clover Clamp"),
            ],
            "Fork Length": [
                ("A", "Standard length 1.7 in. (44 mm)"),
                ("H", "Standard length flange 4.0 in. (102 mm)"),
                ("E", "Extended, customer specified length in tenths of inches"),
                ("M", "Extended, customer specified length in millimeters"),
            ],
            "Fork Length Specification": [
                ("0000", "Factory default length"),
            ],
            "Surface Finish": [
                ("1", "Standard surface finish"),
                ("2", "Hand polished (Ra < 0.4 μm)"),
            ],
            "Product Certificates": [
                ("NA", "No hazardous locations certifications (safe area use only)"),
                ("G5", "USA Ordinary Locations (unclassified, safe area)"),
                ("G6", "Canada Ordinary Locations (unclassified, safe area)"),
                ("E1", "ATEX Flameproof"),
                ("E5", "USA Explosion-proof"),
                ("E6", "Canada Explosion-proof"),
                ("I1", "ATEX Intrinsic Safety"),
                ("I5", "USA Intrinsic Safety"),
                ("I6", "Canada Intrinsic Safety"),
            ],
            "Additional Options": [
                ("Q4", "Certificate of functional test"),
                ("Q8", "Material traceability certification per EN 10204 3.1"),
                ("Q15", "NACE® MR0175 / ISO 15156"),
                ("Q25", "NACE MR0103"),
                ("QS", "Prior-use certificate of FMEDA Data"),
                ("QT", "Safety certified to IEC61508"),
                ("Q73", "Certificate of liquid penetrant inspection"),
                ("Q76", "Positive material identification certificate of conformance"),
                ("P1", "Hydrostatic testing with certificate"),
                ("LD", "Low density liquids – minimum density is 31.2 lb/ft³ (500 kg/m³)"),
                ("WR3", "3-year limited warranty"),
                ("WR5", "5-year limited warranty"),
                ("R2259", "Overfill approval option (TÜV-tested and approved for overfill protection)"),
            ],
        }
        
        self.selections = {}
        self.combo_vars = {}
        self.show_hyphens = False
        self.cert_checkboxes = {}
        self.fork_length_vars = {}
        
        # Create main frame with scrollbar
        self.create_scrollable_frame()
        
        # Create display frame for model number
        self.create_display_frame()
    
    def create_scrollable_frame(self):
        # Frame for scrollbar and canvas
        scroll_frame = ttk.Frame(self.root)
        scroll_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Canvas and scrollbar
        canvas = tk.Canvas(scroll_frame, bg="white")
        scrollbar = ttk.Scrollbar(scroll_frame, orient="vertical", command=canvas.yview)
        
        # Scrollable frame inside canvas
        self.inner_frame = ttk.Frame(canvas, padding=10)
        self.inner_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        # Configure the inner frame to expand
        self.inner_frame.pack(fill=tk.BOTH, expand=True)
        
        canvas.create_window((0, 0), window=self.inner_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Bind canvas resize to update inner frame width
        def configure_inner_frame(event):
            canvas_width = event.width
            canvas.itemconfig(canvas.find_all()[0], width=canvas_width)
        
        canvas.bind('<Configure>', configure_inner_frame)
        
        # Mousewheel scrolling
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        
        canvas.pack(side="left", fill=tk.BOTH, expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Create comboboxes for each field
        for field_name, options in self.data.items():
            if field_name == "Product Certificates" or field_name == "Additional Options":
                self.create_certificates_row(field_name, options)
            elif field_name == "Fork Length Specification":
                self.create_fork_length_row(field_name, options)
            else:
                self.create_combobox_row(field_name, options)
    
    def create_combobox_row(self, field_name, options):
        # Frame for label and combobox
        row_frame = ttk.Frame(self.inner_frame)
        row_frame.pack(fill="x", pady=8)
        
        # Configure grid weights for proper column sizing
        row_frame.columnconfigure(0, weight=0)  # Label column - fixed width
        row_frame.columnconfigure(1, weight=1)   # Combobox column - expandable
        
        # Label with fixed width and proper padding
        label = ttk.Label(row_frame, text=field_name + ":", width=30, anchor="w")
        label.grid(row=0, column=0, sticky="w", padx=(5, 10), pady=2)
        
        # Prepare dropdown values
        dropdown_values = [f"{code} - {desc}" for code, desc in options]
        
        # Combobox that expands to fill available width with minimum width
        var = tk.StringVar()
        combo = ttk.Combobox(
            row_frame,
            textvariable=var,
            values=dropdown_values,
            width=60,  # Set a good minimum width
            state="readonly"
        )
        combo.grid(row=0, column=1, sticky="ew", padx=(0, 5), pady=2)
        
        # Store variable and bind to update
        self.combo_vars[field_name] = var
        var.trace("w", self.update_model_number)
        
        # Store reference for extracting code
        self.selections[field_name] = {
            "var": var,
            "options": options
        }
    
    def create_certificates_row(self, field_name, options):
        # Frame for label and checkboxes
        row_frame = ttk.Frame(self.inner_frame)
        row_frame.pack(fill="x", pady=8)
        
        # Configure grid weights for proper column sizing
        row_frame.columnconfigure(0, weight=0)  # Label column - fixed width
        row_frame.columnconfigure(1, weight=1)   # Checkboxes column - expandable
        
        # Label with fixed width and proper padding
        label = ttk.Label(row_frame, text=field_name + ":", width=30, anchor="w")
        label.grid(row=0, column=0, sticky="w", padx=(5, 10), pady=2)
        
        # Frame for checkboxes
        checkbox_frame = ttk.Frame(row_frame)
        checkbox_frame.grid(row=0, column=1, sticky="ew", padx=(0, 5), pady=2)
        
        # Create checkboxes in a grid layout
        self.cert_checkboxes[field_name] = {}
        for i, (code, desc) in enumerate(options):
            var = tk.BooleanVar()
            checkbox = ttk.Checkbutton(
                checkbox_frame,
                text=f"{code} - {desc}",
                variable=var,
                command=self.update_model_number
            )
            # Arrange in 2 columns
            row = i // 2
            col = i % 2
            checkbox.grid(row=row, column=col, sticky="w", padx=5, pady=2)
            
            self.cert_checkboxes[field_name][code] = var
        
        # Store reference for extracting codes
        self.selections[field_name] = {
            "checkboxes": self.cert_checkboxes[field_name],
            "options": options
        }
    
    def create_fork_length_row(self, field_name, options):
        # Frame for label and input controls
        row_frame = ttk.Frame(self.inner_frame)
        row_frame.pack(fill="x", pady=8)
        
        # Configure grid weights for proper column sizing
        row_frame.columnconfigure(0, weight=0)  # Label column - fixed width
        row_frame.columnconfigure(1, weight=1)   # Input controls column - expandable
        
        # Label with fixed width and proper padding
        label = ttk.Label(row_frame, text=field_name + ":", width=30, anchor="w")
        label.grid(row=0, column=0, sticky="w", padx=(5, 10), pady=2)
        
        # Frame for input controls
        input_frame = ttk.Frame(row_frame)
        input_frame.grid(row=0, column=1, sticky="ew", padx=(0, 5), pady=2)
        
        # Input box for length value
        length_var = tk.StringVar()
        length_entry = ttk.Entry(
            input_frame,
            textvariable=length_var,
            width=15,
            justify="center"
        )
        length_entry.pack(side="left", padx=(0, 5))
        
        # Dropdown for units (Inches/Millimeters)
        unit_var = tk.StringVar()
        unit_combo = ttk.Combobox(
            input_frame,
            textvariable=unit_var,
            values=["Inches", "Millimeters"],
            width=12,
            state="readonly"
        )
        unit_combo.pack(side="left", padx=(0, 5))
        
        # Bind to update model number
        length_var.trace("w", self.update_model_number)
        unit_var.trace("w", self.update_model_number)
        
        # Bind Fork Length selection to control units dropdown
        fork_length_var = self.combo_vars.get("Fork Length")
        if fork_length_var:
            fork_length_var.trace("w", lambda *args: self.update_fork_length_units(fork_length_var.get(), unit_combo, length_entry))
        
        # Store variables
        self.fork_length_vars[field_name] = {
            "length": length_var,
            "unit": unit_var
        }
        
        # Store reference for extracting code
        self.selections[field_name] = {
            "length_var": length_var,
            "unit_var": unit_var,
            "options": options
        }
    
    def update_fork_length_units(self, fork_length_selection, unit_combo, length_entry):
        """Update units dropdown and length entry based on Fork Length selection"""
        if not fork_length_selection:
            return
            
        # Extract the code from the selection
        fork_code = fork_length_selection.split(" - ")[0]
        
        if fork_code in ["A", "H"]:
            # Standard lengths - disable input and clear
            length_entry.config(state="disabled")
            unit_combo.config(state="disabled")
            length_entry.delete(0, tk.END)
            unit_combo.set("")
        elif fork_code == "E":
            # Extended in inches - enable input and set units
            length_entry.config(state="normal")
            unit_combo.config(state="readonly")
            unit_combo.set("Inches")
        elif fork_code == "M":
            # Extended in millimeters - enable input and set units
            length_entry.config(state="normal")
            unit_combo.config(state="readonly")
            unit_combo.set("Millimeters")
    
    def create_display_frame(self):
        display_frame = ttk.LabelFrame(self.root, text="Generated Model Number", padding=10)
        display_frame.pack(fill="x", padx=10, pady=10)
        
        # Model number display (now selectable)
        self.model_display = tk.Entry(
            display_frame,
            font=tkFont.Font(family="Courier", size=14, weight="bold"),
            fg="blue",
            state="readonly",
            justify="center",
            width=30
        )
        self.model_display.pack(pady=5)
        
        # Button frame
        button_frame = ttk.Frame(display_frame)
        button_frame.pack(pady=5)
        
        # Copy button
        copy_button = ttk.Button(
            button_frame,
            text="Copy to Clipboard",
            command=self.copy_model_number
        )
        copy_button.pack(side="left", padx=5)
        
        # Toggle hyphens button
        self.hyphen_button = ttk.Button(
            button_frame,
            text="Add Hyphens",
            command=self.toggle_hyphens
        )
        self.hyphen_button.pack(side="left", padx=5)
    
    def update_model_number(self, *args):
        model_number = "2130"
        
        for field_name in self.data.keys():
            if field_name == "Model":
                continue
            
            selection_data = self.selections[field_name]
            
            if field_name == "Product Certificates" or field_name == "Additional Options":
                # Handle checkboxes for certificates and additional options
                selected_options = []
                for code, var in selection_data["checkboxes"].items():
                    if var.get():
                        selected_options.append(code)
                
                # Add all selected options
                for option in selected_options:
                    model_number += option
            elif field_name == "Fork Length Specification":
                # Handle fork length input - only process if E or M is selected in Fork Length
                fork_length_selection = self.selections.get("Fork Length", {}).get("var", tk.StringVar()).get()
                if fork_length_selection:
                    fork_code = fork_length_selection.split(" - ")[0]
                    
                    if fork_code in ["A", "H"]:
                        # Standard lengths - use factory default
                        model_number += "0000"
                    elif fork_code in ["E", "M"]:
                        # Extended lengths - use user input
                        length_value = selection_data["length_var"].get()
                        unit_value = selection_data["unit_var"].get()
                        
                        if length_value and unit_value:
                            try:
                                if fork_code == "E" and unit_value == "Inches":
                                    # Convert to tenths of inches (multiply by 10)
                                    length_in_tenths = int(float(length_value) * 10)
                                    # Format as 4-digit code (e.g., 118.1 inches = 1181)
                                    code = f"{length_in_tenths:04d}"
                                    model_number += code
                                elif fork_code == "M" and unit_value == "Millimeters":
                                    # Use as-is for millimeters (e.g., 3000 mm = 3000)
                                    length_mm = int(float(length_value))
                                    code = f"{length_mm:04d}"
                                    model_number += code
                            except (ValueError, TypeError):
                                # If invalid input, skip this field
                                pass
            else:
                # Handle regular combobox
                current_value = selection_data["var"].get()
                
                if current_value:
                    # Extract code from "CODE - Description" format
                    code = current_value.split(" - ")[0]
                    model_number += code
        
        # Apply hyphen formatting if enabled
        if self.show_hyphens and len(model_number) > 4:
            # Add hyphens after the base model (2130) and at strategic points
            formatted_number = model_number[:4]  # Start with "2130"
            remaining = model_number[4:]
            
            # Add hyphens at logical break points (every 2-3 characters)
            i = 0
            while i < len(remaining):
                if i > 0 and i % 2 == 0:
                    formatted_number += "-"
                formatted_number += remaining[i]
                i += 1
            
            model_number = formatted_number
        
        # Update the display
        self.model_display.config(state="normal")
        self.model_display.delete(0, tk.END)
        self.model_display.insert(0, model_number)
        self.model_display.config(state="readonly")
    
    def copy_model_number(self):
        """Copy the current model number to clipboard"""
        model_number = self.model_display.get()
        pyperclip.copy(model_number)
        
        # Provide visual feedback
        original_text = self.model_display.get()
        self.model_display.config(state="normal")
        self.model_display.delete(0, tk.END)
        self.model_display.insert(0, "Copied to clipboard!")
        self.model_display.config(state="readonly")
        
        # Reset after 1 second
        self.root.after(1000, lambda: self.reset_copy_feedback(original_text))
    
    def reset_copy_feedback(self, original_text):
        """Reset the display after copy feedback"""
        self.model_display.config(state="normal")
        self.model_display.delete(0, tk.END)
        self.model_display.insert(0, original_text)
        self.model_display.config(state="readonly")
    
    def toggle_hyphens(self):
        """Toggle hyphen display in model number"""
        self.show_hyphens = not self.show_hyphens
        
        # Update button text
        if self.show_hyphens:
            self.hyphen_button.config(text="Remove Hyphens")
        else:
            self.hyphen_button.config(text="Add Hyphens")
        
        # Refresh the model number display
        self.update_model_number()

if __name__ == "__main__":
    root = tk.Tk()
    app = Rosemount2130Builder(root)
    root.mainloop()