import tkinter as tk
from tkinter import ttk, messagebox
from typing import Dict, List

class AshcroftModelBuilder:
    def __init__(self, root):
        self.root = root
        self.root.title("Ashcroft Flanged Diaphragm Seal Model Builder")
        self.root.geometry("900x700")
        
        # Variables for selections
        self.vars = {
            'process_size': tk.StringVar(),
            'diaphragm_type': tk.StringVar(),
            'bottom_housing': tk.StringVar(),
            'diaphragm_material': tk.StringVar(),
            'bottom_material': tk.StringVar(),
            'instrument_connection': tk.StringVar(),
            'fill_fluid': tk.StringVar(),
            'flange_rating': tk.StringVar(),
            'flange_type': tk.StringVar()
        }
        
        # Optional features
        self.optional_vars = {
            'flushing_port': tk.StringVar(value=''),
            'top_housing': tk.StringVar(value=''),
            'assembly': tk.StringVar(value=''),
            'other': tk.StringVar(value='')
        }
        
        self.create_widgets()
        self.update_availability()
    
    def create_widgets(self):
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        self.root.rowconfigure(1, weight=0)  # Fixed bottom section
        
        # Create fixed bottom frame for model number (always visible)
        bottom_frame = ttk.Frame(self.root, padding="10")
        bottom_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.S))
        bottom_frame.columnconfigure(0, weight=1)
        
        # Model Number Display (always visible)
        ttk.Separator(bottom_frame, orient='horizontal').grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        ttk.Label(bottom_frame, text="Generated Model Number:", font=('Arial', 12, 'bold')).grid(row=1, column=0, sticky=tk.W, pady=5)
        
        self.model_display = tk.Text(bottom_frame, height=3, width=70, font=('Courier', 12, 'bold'))
        self.model_display.grid(row=2, column=0, sticky=(tk.W, tk.E), padx=20, pady=5)
        
        # Buttons
        button_frame = ttk.Frame(bottom_frame)
        button_frame.grid(row=3, column=0, pady=10)
        ttk.Button(button_frame, text="Copy to Clipboard", command=self.copy_to_clipboard).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Reset", command=self.reset_form).pack(side=tk.LEFT, padx=5)
        
        # Create scrollable frame for form content
        canvas = tk.Canvas(self.root)
        scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas, padding="10")
        
        # Create canvas window and store reference
        canvas_window = canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        
        def configure_scroll_region(event=None):
            canvas.configure(scrollregion=canvas.bbox("all"))
            # Update canvas window width to match canvas width
            canvas_width = event.width if event else canvas.winfo_width()
            if canvas_width > 1:  # Only update if canvas has been rendered
                canvas.itemconfig(canvas_window, width=canvas_width)
        
        scrollable_frame.bind("<Configure>", configure_scroll_region)
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Update canvas width when root window is resized
        def on_canvas_configure(event):
            canvas_width = event.width
            canvas.itemconfig(canvas_window, width=canvas_width)
        canvas.bind('<Configure>', on_canvas_configure)
        
        canvas.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        scrollable_frame.columnconfigure(1, weight=1)
        
        row = 0
        
        # Title
        title = ttk.Label(scrollable_frame, text="Ashcroft 102/103, 202/203, 302/303 Flanged Diaphragm Seal", 
                         font=('Arial', 14, 'bold'))
        title.grid(row=row, column=0, columnspan=2, pady=(0, 15))
        row += 1
        
        # Process Connection Size
        ttk.Label(scrollable_frame, text="Process Connection Size:", font=('Arial', 10, 'bold')).grid(row=row, column=0, sticky=tk.W, pady=5)
        row += 1
        sizes = [('½ NPS', '50'), ('¾ NPS', '75'), ('1 NPS', '10'), 
                ('1½ NPS', '15'), ('2 NPS', '20'), ('3 NPS', '30')]
        for text, value in sizes:
            ttk.Radiobutton(scrollable_frame, text=text, variable=self.vars['process_size'], 
                           value=value, command=self.update_model).grid(row=row, column=0, sticky=tk.W, padx=20)
            row += 1
        
        # Diaphragm Type
        ttk.Label(scrollable_frame, text="Diaphragm Type:", font=('Arial', 10, 'bold')).grid(row=row, column=0, sticky=tk.W, pady=5)
        row += 1
        types = [('100 Series - Metallic capsule threaded', '1'),
                ('200 Series - Welded/bonded', '2'),
                ('300 Series - Elastomeric clamped', '3')]
        for text, value in types:
            ttk.Radiobutton(scrollable_frame, text=text, variable=self.vars['diaphragm_type'], 
                           value=value, command=self.update_model).grid(row=row, column=0, sticky=tk.W, padx=20)
            row += 1
        
        # Bottom Housing Options
        ttk.Label(scrollable_frame, text="Bottom Housing Options:", font=('Arial', 10, 'bold')).grid(row=row, column=0, sticky=tk.W, pady=5)
        row += 1
        housing = [('Flanged without flushing', '02'),
                  ('Flanged with flushing connection', '03')]
        for text, value in housing:
            ttk.Radiobutton(scrollable_frame, text=text, variable=self.vars['bottom_housing'], 
                           value=value, command=self.update_model).grid(row=row, column=0, sticky=tk.W, padx=20)
            row += 1
        
        # Diaphragm Material
        ttk.Label(scrollable_frame, text="Diaphragm Material:", font=('Arial', 10, 'bold')).grid(row=row, column=0, sticky=tk.W, pady=5)
        row += 1
        ttk.Combobox(scrollable_frame, textvariable=self.vars['diaphragm_material'], 
                    values=self.get_diaphragm_materials(), width=40, 
                    state='readonly').grid(row=row, column=0, columnspan=2, sticky=tk.W, padx=20, pady=2)
        self.vars['diaphragm_material'].trace('w', lambda *args: self.update_model())
        row += 1
        
        # Bottom Housing Material
        ttk.Label(scrollable_frame, text="Bottom Housing Material:", font=('Arial', 10, 'bold')).grid(row=row, column=0, sticky=tk.W, pady=5)
        row += 1
        ttk.Combobox(scrollable_frame, textvariable=self.vars['bottom_material'], 
                    values=self.get_bottom_materials(), width=40, 
                    state='readonly').grid(row=row, column=0, columnspan=2, sticky=tk.W, padx=20, pady=2)
        self.vars['bottom_material'].trace('w', lambda *args: self.update_model())
        row += 1
        
        # Instrument Connection Size
        ttk.Label(scrollable_frame, text="Instrument Connection:", font=('Arial', 10, 'bold')).grid(row=row, column=0, sticky=tk.W, pady=5)
        row += 1
        connections = [('¼ NPT Female', '02T'), ('½ NPT Female', '04T')]
        for text, value in connections:
            ttk.Radiobutton(scrollable_frame, text=text, variable=self.vars['instrument_connection'], 
                           value=value, command=self.update_model).grid(row=row, column=0, sticky=tk.W, padx=20)
            row += 1
        
        # Fill Fluid
        ttk.Label(scrollable_frame, text="Fill Fluid:", font=('Arial', 10, 'bold')).grid(row=row, column=0, sticky=tk.W, pady=5)
        row += 1
        ttk.Combobox(scrollable_frame, textvariable=self.vars['fill_fluid'], 
                    values=self.get_fill_fluids(), width=40, 
                    state='readonly').grid(row=row, column=0, columnspan=2, sticky=tk.W, padx=20, pady=2)
        self.vars['fill_fluid'].trace('w', lambda *args: self.update_model())
        row += 1
        
        # Flange Rating
        ttk.Label(scrollable_frame, text="Flange Rating:", font=('Arial', 10, 'bold')).grid(row=row, column=0, sticky=tk.W, pady=5)
        row += 1
        ttk.Combobox(scrollable_frame, textvariable=self.vars['flange_rating'], 
                    values=['150', '300', '600', '900', '1500', '2500'], width=40, 
                    state='readonly').grid(row=row, column=0, columnspan=2, sticky=tk.W, padx=20, pady=2)
        self.vars['flange_rating'].trace('w', lambda *args: self.update_model())
        row += 1
        
        # Flange Type
        ttk.Label(scrollable_frame, text="Flange Type:", font=('Arial', 10, 'bold')).grid(row=row, column=0, sticky=tk.W, pady=5)
        row += 1
        types = [('Raised Face', 'RF'), ('Ring Joint', 'RJ'), ('Flat Face', 'FF')]
        for text, value in types:
            ttk.Radiobutton(scrollable_frame, text=text, variable=self.vars['flange_type'], 
                           value=value, command=self.update_model).grid(row=row, column=0, sticky=tk.W, padx=20)
            row += 1
        
        # Optional Features
        ttk.Separator(scrollable_frame, orient='horizontal').grid(row=row, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=10)
        row += 1
        ttk.Label(scrollable_frame, text="Optional Features:", font=('Arial', 10, 'bold')).grid(row=row, column=0, sticky=tk.W, pady=5)
        row += 1
        
        # Flushing Port (only if 03 selected)
        ttk.Label(scrollable_frame, text="Flushing Port:").grid(row=row, column=0, sticky=tk.W, padx=20)
        ttk.Combobox(scrollable_frame, textvariable=self.optional_vars['flushing_port'], 
                    values=['', 'AW - Single ½" flushing', 'DB - Dual ½" flushing', 
                           'DK - Dual ¼" flushing', 'PU - Pipe plug'], width=35, 
                    state='readonly').grid(row=row, column=1, sticky=tk.W, pady=2)
        self.optional_vars['flushing_port'].trace('w', lambda *args: self.update_model())
        row += 1
        
        # Assembly Options
        ttk.Label(scrollable_frame, text="Assembly/Hardware:").grid(row=row, column=0, sticky=tk.W, padx=20)
        ttk.Combobox(scrollable_frame, textvariable=self.optional_vars['assembly'], 
                    values=['', 'SE - SS flange rings and bolts', 'SB - SS clamping bolts', 
                           'LD - SS locking device', 'NH - SS instrument tag', 'DU - Instrument welded'], 
                    width=35, state='readonly').grid(row=row, column=1, sticky=tk.W, pady=2)
        self.optional_vars['assembly'].trace('w', lambda *args: self.update_model())
        row += 1
        
        # Other Options
        ttk.Label(scrollable_frame, text="Other:").grid(row=row, column=0, sticky=tk.W, padx=20)
        ttk.Combobox(scrollable_frame, textvariable=self.optional_vars['other'], 
                    values=['', 'CD-5 - NACE compliance', 'MQ - Positive material ID', 
                           '6B - Cleaned for oxygen service'], width=35, 
                    state='readonly').grid(row=row, column=1, sticky=tk.W, pady=2)
        self.optional_vars['other'].trace('w', lambda *args: self.update_model())
        
        # Bind mousewheel to canvas for scrolling
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)
    
    def get_diaphragm_materials(self) -> List[str]:
        return [
            'C - 304L SS', 'S - 316L SS', 'F - 904L SS', 'D - Carpenter 20',
            'W - Gold Plated 316L SS', 'R - Halar-coated Monel', 'G - Hastelloy B',
            'J - Hastelloy C-22', 'H - Hastelloy C-276', 'P - Monel 400',
            'N - Nickel', 'U - Tantalum', 'Ti - Titanium', 'K - Kalrez',
            'T - PTFE', 'Y - Viton'
        ]
    
    def get_bottom_materials(self) -> List[str]:
        return [
            'C - 304L SS', 'S - 316L SS', 'Q - 321 SS', 'E - 347 SS',
            'F - 904L SS', 'D - Carpenter 20', 'Z - Duplex 2205',
            'BH - Halar-coated Monel', 'G - Hastelloy B', 'J - Hastelloy C-22',
            'H - Hastelloy C-276', 'W - Inconel 625', 'M - Monel 400',
            'N - Nickel', 'T - PTFE', 'V - PVC', 'KY - PVDF (Kynar)',
            'B - Steel', 'SU - Tantalum-clad 316L SS', 'Ti - Titanium'
        ]
    
    def get_fill_fluids(self) -> List[str]:
        return [
            'CG - Glycerin (food grade)', 'CK - 50 cSt Silicone', 'DJ - 10 cSt Silicone',
            'CF - Halocarbon 4.2', 'HA - Slytherm 800', 'CC - Syltherm XLT',
            'KF - Calflo AF', 'MY - Mineral Oil', 'NM - Neobee M-20 (food grade)',
            'CZ - Silicone (food grade)', 'FJ - Distilled Water', 'GH - 50/50 Glycerin/Water',
            'CV - Propylene Glycol', 'FK - Ethylene Glycol', 'CT - 50/50 Ethylene Glycol/Water',
            'GR - 80/20 Glycerin/Water', 'PY - 95/5 Water/Propylene Glycol'
        ]
    
    def update_availability(self):
        # Update which options are available based on selections
        pass
    
    def update_model(self):
        # Build model number
        parts = []
        
        # Process size
        if self.vars['process_size'].get():
            parts.append(self.vars['process_size'].get())
        
        # Diaphragm type
        if self.vars['diaphragm_type'].get():
            parts.append(self.vars['diaphragm_type'].get())
        
        # Bottom housing
        if self.vars['bottom_housing'].get():
            parts.append(self.vars['bottom_housing'].get())
        
        # Diaphragm material (extract code)
        if self.vars['diaphragm_material'].get():
            code = self.vars['diaphragm_material'].get().split(' - ')[0]
            parts.append(code)
        
        # Bottom material (extract code)
        if self.vars['bottom_material'].get():
            code = self.vars['bottom_material'].get().split(' - ')[0]
            parts.append(code)
        
        # Instrument connection
        if self.vars['instrument_connection'].get():
            parts.append(self.vars['instrument_connection'].get())
        
        # Options section
        options = []
        
        # Fill fluid
        if self.vars['fill_fluid'].get():
            options.append(self.vars['fill_fluid'].get().split(' - ')[0])
        
        # Optional features
        for key, var in self.optional_vars.items():
            if var.get():
                options.append(var.get().split(' - ')[0])
        
        # Add options with -X prefix
        if options:
            parts.append('-X' + ''.join(options))
        
        # Flange rating
        if self.vars['flange_rating'].get():
            parts.append(self.vars['flange_rating'].get())
        
        # Flange type
        if self.vars['flange_type'].get():
            parts.append(self.vars['flange_type'].get())
        
        # Display model number
        model = ' '.join(parts)
        self.model_display.delete('1.0', tk.END)
        self.model_display.insert('1.0', model)
    
    def copy_to_clipboard(self):
        model = self.model_display.get('1.0', tk.END).strip()
        if model:
            self.root.clipboard_clear()
            self.root.clipboard_append(model)
            messagebox.showinfo("Success", "Model number copied to clipboard!")
        else:
            messagebox.showwarning("Warning", "Please complete the configuration first.")
    
    def reset_form(self):
        for var in self.vars.values():
            var.set('')
        for var in self.optional_vars.values():
            var.set('')
        self.model_display.delete('1.0', tk.END)

if __name__ == "__main__":
    root = tk.Tk()
    app = AshcroftModelBuilder(root)
    root.mainloop()