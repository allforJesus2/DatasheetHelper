import tkinter as tk
from tkinter import ttk

# Data extracted from the Rosemount 3051 PDF
# This dictionary holds the options for building the model number.
# The 'order' key helps separate core model number parts from appended options.
MODEL_DATA = {
    "core": [
        "Transmitter Type", "Pressure Range", "Output", "Materials of Construction",
        "Isolating Diaphragm", "O-ring", "Fill Fluid", "Housing Material"
    ],
    "append": [
        "Alternate Flange", "Mounting Bracket", "Hazardous Locations Certification",
        "Bolting", "Display", "Terminal Block", "Transmitter Version"
    ],
    "options": {
        "Transmitter Type": {
            "CD": {"text": "CD: Differential Pressure"},
            "CG": {"text": "CG: Gage Pressure"},
            "CA": {"text": "CA: Absolute Pressure"},
        },
        "Pressure Range": {
            "0": {"text": "0: -3 to 3 inH2O / 0.1 inH2O", "depends": ["CD"]},
            "1": {"text": "1: -25 to 25 inH2O (CD/CG) or 0 to 30 psia (CA)", "depends": ["CD", "CG", "CA"]},
            "2": {"text": "2: -250 to 250 inH2O (CD/CG) or 0 to 150 psia (CA)", "depends": ["CD", "CG", "CA"]},
            "3": {"text": "3: -1000 to 1000 inH2O (CD/CG) or 0 to 800 psia (CA)", "depends": ["CD", "CG", "CA"]},
            "4": {"text": "4: -300 to 300 psi (CD/CG) or 0 to 4000 psia (CA)", "depends": ["CD", "CG", "CA"]},
            "5": {"text": "5: -2000 to 2000 psi", "depends": ["CD", "CG"]},
        },
        "Output": {
            "A": {"text": "A: 4-20 mA with Digital Signal Based on HART Protocol"}
        },
        "Materials of Construction": {
            "2": {"text": "2: Coplanar, SST Flange, SST Drain/Vent"},
            "3": {"text": "3: Coplanar, Hastelloy C-276 Flange & Drain/Vent"},
            "4": {"text": "4: Coplanar, Monel Flange & Drain/Vent"},
            "5": {"text": "5: Coplanar, Plated CS Flange, SST Drain/Vent"},
            "7": {"text": "7: Coplanar, SST Flange, Hastelloy C-276 Drain/Vent"},
            "8": {"text": "8: Coplanar, Hastelloy C-276 Flange, Plated CS"},
            "0": {"text": "0: Alternate Flange (See option below)"},
        },
        "Alternate Flange": {
            "": {"text": "None"},
            "H2": {"text": "H2: Traditional Flange, 316 SST, SST Drain/Vent"},
            "H3": {"text": "H3: Traditional Flange, Alloy C, Hastelloy C276 Drain/Vent"},
            "H4": {"text": "H4: Traditional Flange, Monel, Monel Drain/Vent"},
            "H7": {"text": "H7: Traditional Flange, 316 SST, Hastelloy C-276 Drain/Vent"},
            "H7": {"text": "H7: Traditional Flange, 316 SST, Hastelloy C276 Drain/Vent"},
            "HJ": {"text": "HJ: DIN Compliant Traditional Flange, SST, 7/16 in. Adapter/Manifold Bolting"},
            "HK": {"text": "HK: DIN Compliant Traditional Flange, SST, 10 mm Adapter/Manifold Bolting"},
            "HL": {"text": "HL: DIN Compliant Traditional Flange, SST, 12mm Adapter/Manifold Bolting (Not available on 3051CD0)"},
            "FA": {"text": "FA: Level Flange, SST, 2 in., ANSI Class 150, Vertical Mount"},
            "FB": {"text": "FB: Level Flange, SST, 2 in., ANSI Class 300, Vertical Mount"},
            "FC": {"text": "FC: Level Flange, SST, 3 in., ANSI Class 150, Vertical Mount"},
            "FD": {"text": "FD: Level Flange, SST, 3 in., ANSI Class 300, Vertical Mount"},
            "FP": {"text": "FP: DIN Level Flange, SST, DN 50, PN 40, Vertical Mount"},
            "FQ": {"text": "FQ: DIN Level Flange, SST, DN 80, PN 40, Vertical Mount"},
            "S5": {"text": "S5: Assemble to Rosemount 305 Integral Manifold (specified separately, see the Rosemount 305 and 306 Integral Manifolds PDS (document number 00813-0100-4733))"},
            "S6": {"text": "S6: Assemble to Rosemount 304 Manifold or Connection System"},
        },
        "Isolating Diaphragm": {
            "2": {"text": "2: 316L SST"},
            "3": {"text": "3: Hastelloy C-276"},
            "4": {"text": "4: Monel"},
            "5": {"text": "5: Tantalum"},
            "6": {"text": "6: Gold-plated Monel"},
            "7": {"text": "7: Gold-plated SST"},
        },
        "O-ring": {
            "A": {"text": "A: Glass-filled PTFE"},
            "B": {"text": "B: Graphite-filled PTFE"},
        },
        "Fill Fluid": {
            "1": {"text": "1: Silicone"},
            "2": {"text": "2: Inert fill (Halocarbon)"},
        },
        "Housing Material": {
            "A": {"text": "A: Aluminum, 1/2-14 NPT"},
            "B": {"text": "B: Aluminum, M20 x 1.5"},
            "D": {"text": "D: Aluminum, G1/2"},
            "J": {"text": "J: SST, 1/2-14 NPT"},
            "K": {"text": "K: SST, M20 x 1.5"},
            "M": {"text": "M: SST, G1/2"},
        },
        "Mounting Bracket": {
            "": {"text": "None"},
            "B4": {"text": "B4: Coplanar Flange Bracket, SST"},
            "B1": {"text": "B1: Traditional Flange Bracket for 2-in. Pipe, CS"},
            "B2": {"text": "B2: Traditional Flange Bracket for Panel, CS"},
            "B3": {"text": "B3: Traditional Flange Flat Bracket for 2-in. Pipe, CS"},
            "B7": {"text": "B7: B1 Bracket with SST Bolts"},
            "B8": {"text": "B8: B2 Bracket with SST Bolts"},
            "B9": {"text": "B9: B3 Bracket with SST Bolts"},
        },
        "Hazardous Locations Certification": {
            "": {"text": "None"},
            "E5": {"text": "E5: FM Explosion-proof"},
            "I5": {"text": "I5: FM Intrinsically safe, non-incendive"},
            "K5": {"text": "K5: FM Combination of E5 and I5"},
            "I1": {"text": "I1: ATEX Intrinsically safe, Dust"},
            "N1": {"text": "N1: ATEX Type n, Dust"},
            "E8": {"text": "E8: ATEX Flameproof, Dust"},
            "C6": {"text": "C6: CSA Explosion-proof, Intrinsically safe"},
            "KD": {"text": "KD: FM, CSA, and ATEX Explosion-proof, Intrinsically Safe"},
        },
        "Bolting": {
            "": {"text": "None"},
            "L4": {"text": "L4: Austenitic 316 SST Bolts"},
            "L5": {"text": "L5: ASTM A 193, Grade B7M Bolts"},
            "L6": {"text": "L6: Monel Bolts"},
            "L8": {"text": "L8: ASTM A 193 Class 2, Grade B8M Bolts"},
        },
        "Display": {
            "": {"text": "None"},
            "M5": {"text": "M5: LCD Display"},
        },
        "Terminal Block": {
            "": {"text": "None"},
            "T1": {"text": "T1: Transient Protection Terminal Block"},
        },
        "Transmitter Version": {
            "TR": {"text": "TR: Transmitter Version 5 (Required)"}
        }
    }
}


class ScrollableFrame(ttk.Frame):
    """A scrollable frame widget for tkinter."""
    def __init__(self, container, *args, **kwargs):
        super().__init__(container, *args, **kwargs)
        canvas = tk.Canvas(self, borderwidth=0, background="#ffffff")
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        self.scrollable_frame = ttk.Frame(canvas, padding=(0,0,15,0)) # Add padding for scrollbar visibility

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        # Bind mouse wheel event to the canvas and the scrollable frame
        canvas.bind_all("<MouseWheel>", _on_mousewheel)

        canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

class App(tk.Tk):
    """Main application class for the model builder."""
    def __init__(self):
        super().__init__()
        self.title("Rosemount 3051 Model Builder")
        self.geometry("700x750")

        self.widgets = {}
        self.all_categories = MODEL_DATA["core"] + MODEL_DATA["append"]
        
        # --- Model Number Display ---
        self.model_frame = ttk.Frame(self, padding="10")
        self.model_frame.pack(side="top", fill="x", pady=5)
        
        ttk.Label(self.model_frame, text="Generated Model:", font=("Helvetica", 12, "bold")).pack(side="left")
        self.model_var = tk.StringVar(value="3051")
        self.model_label = ttk.Label(self.model_frame, textvariable=self.model_var, font=("Helvetica", 12, "bold"), foreground="blue", wraplength=550)
        self.model_label.pack(side="left", padx=10)

        # --- Scrollable Area for Options ---
        scrollable_container = ScrollableFrame(self)
        scrollable_container.pack(fill="both", expand=True, padx=10, pady=10)
        self.main_frame = scrollable_container.scrollable_frame

        # --- Create Dropdowns and Labels ---
        for i, category in enumerate(self.all_categories):
            options = MODEL_DATA['options'].get(category, {})
            option_frame = ttk.LabelFrame(self.main_frame, text=category, padding="10")
            option_frame.grid(row=i, column=0, padx=10, pady=10, sticky="ew")
            self.main_frame.grid_columnconfigure(0, weight=1)

            combo = ttk.Combobox(option_frame, state="readonly", width=60)
            combo.pack(fill="x")
            combo.bind("<<ComboboxSelected>>", self.update_model_number)
            
            self.widgets[category] = {
                'combo': combo,
                'options': options,
                'frame': option_frame
            }

        self.setup_initial_dropdowns()
        self.update_model_number()

    def get_code_from_text(self, category, text):
        """Finds the option code corresponding to the selected display text."""
        if not text: return ""
        for code, details in self.widgets[category]['options'].items():
            if details['text'] == text:
                return code
        return ""

    def setup_initial_dropdowns(self):
        """Populates the dropdowns with their initial values."""
        for category, widget_info in self.widgets.items():
            options = widget_info['options']
            widget_info['combo']['values'] = [opt['text'] for opt in options.values()]
            if widget_info['combo']['values']:
                 widget_info['combo'].current(0)
        
        # Transmitter Version is required
        self.widgets["Transmitter Version"]['combo'].set("TR: Transmitter Version 5 (Required)")
        
        self.update_dependent_dropdowns()

    def update_dependent_dropdowns(self):
        """Updates dropdown options based on selections."""
        # 1. Transmitter Type -> Pressure Range dependency
        selected_transmitter_text = self.widgets["Transmitter Type"]['combo'].get()
        selected_transmitter_code = self.get_code_from_text("Transmitter Type", selected_transmitter_text)
        
        range_widget = self.widgets["Pressure Range"]
        valid_ranges = {
            code: details for code, details in range_widget['options'].items()
            if 'depends' not in details or selected_transmitter_code in details['depends']
        }
        
        current_range_text = range_widget['combo'].get()
        range_widget['combo']['values'] = [opt['text'] for opt in valid_ranges.values()]
        
        if current_range_text not in range_widget['combo']['values'] and range_widget['combo']['values']:
             range_widget['combo'].current(0)

        # 2. Materials of Construction -> Alternate Flange dependency
        materials_text = self.widgets["Materials of Construction"]['combo'].get()
        materials_code = self.get_code_from_text("Materials of Construction", materials_text)
        
        alt_flange_combo = self.widgets["Alternate Flange"]['combo']
        if materials_code == '0':
            alt_flange_combo.config(state="readonly")
        else:
            alt_flange_combo.config(state="disabled")
            alt_flange_combo.set("")


    def update_model_number(self, event=None):
        """Constructs the model number string from selected dropdown values."""
        if event:
            self.update_dependent_dropdowns()
        
        # Build the core part of the model number
        core_parts = ["3051"]
        for category in MODEL_DATA["core"]:
            widget_info = self.widgets[category]
            selected_text = widget_info['combo'].get()
            core_parts.append(self.get_code_from_text(category, selected_text))
        
        core_model_string = "".join(core_parts)
        
        # Build the appended options part
        append_parts = []
        for category in MODEL_DATA["append"]:
            widget_info = self.widgets[category]
            
            # Skip disabled alternate flange
            if category == "Alternate Flange" and widget_info['combo']['state'] == 'disabled':
                continue
                
            selected_text = widget_info['combo'].get()
            code = self.get_code_from_text(category, selected_text)
            if code:
                append_parts.append(code)
                
        final_model_string = core_model_string
        if append_parts:
            final_model_string += " " + " ".join(append_parts)
            
        self.model_var.set(final_model_string)


if __name__ == "__main__":
    app = App()
    app.mainloop()

