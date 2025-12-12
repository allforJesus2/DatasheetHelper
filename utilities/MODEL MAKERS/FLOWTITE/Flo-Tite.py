import tkinter as tk
from tkinter import ttk, messagebox
import pyperclip

class FlotiteValveConfigurator:
    def __init__(self, root):
        self.root = root
        self.root.title("Flo-Tite Valve Configurator")
        self.root.geometry("800x900")
        self.root.configure(bg='#f8fafc')
        
        # Configuration state
        self.config = {
            'model': 'F150',
            'bodyMaterial': 'SS',
            'seat': 'F',
            'stemSeal': 'F',
            'bodySeal': 'G',
            'oRings': 'V',
            'size': '50',
            'operator': 'L',
            'special': ''
        }
        
        # Options data
        self.options = {
            'model': [
                {'code': 'F150', 'desc': 'Flanged End Full Bore - Class 150'},
                {'code': 'F300', 'desc': 'Flanged End Full Bore - Class 300'},
                {'code': 'F600', 'desc': 'Flanged End Full Bore - Class 600'},
                {'code': 'SF150', 'desc': 'Flanged End Standard Bore - Class 150'},
                {'code': 'SF300', 'desc': 'Flanged End Standard Bore - Class 300'},
                {'code': 'SF600', 'desc': 'Flanged End Standard Bore - Class 600'},
                {'code': 'RF15', 'desc': 'Flanged End Standard Bore - Class 150 (2PC)'},
                {'code': 'RF30', 'desc': 'Flanged End Standard Bore - Class 300 (2PC)'}
            ],
            'bodyMaterial': [
                {'code': 'SS', 'desc': '316 Stainless Steel'},
                {'code': 'CS', 'desc': 'Carbon Steel (WCB)'},
                {'code': 'A2', 'desc': 'Alloy 20'},
                {'code': 'DP', 'desc': 'Duplex Stainless Steel'}
            ],
            'seat': [
                {'code': 'F', 'desc': 'TFM (Modified PTFE)'},
                {'code': 'Y', 'desc': 'CTFM (Carbon Filled TFM)'},
                {'code': 'T', 'desc': 'PTFE (Teflon)'},
                {'code': 'R', 'desc': 'RPTFE (Reinforced PTFE)'},
                {'code': 'S', 'desc': '50/50 (50% PTFE + SS316)'},
                {'code': 'U', 'desc': 'UHMWPE (Ultra High Molecular Weight PE)'},
                {'code': 'P', 'desc': 'PEEK (Polyetheretherketone)'},
                {'code': 'G', 'desc': 'Graphite'},
                {'code': 'C', 'desc': 'Cavity Filled'},
                {'code': 'M', 'desc': 'Metal Seat'},
                {'code': 'K', 'desc': 'Kel-F'}
            ],
            'stemSeal': [
                {'code': 'F', 'desc': 'TFM Packing'},
                {'code': 'Y', 'desc': 'CTFM (Carbon Filled TFM)'},
                {'code': 'X', 'desc': 'RTFM (Reinforced TFM)'},
                {'code': 'T', 'desc': 'RPTFE (Reinforced PTFE)'},
                {'code': 'R', 'desc': 'PTFE Packing'},
                {'code': 'S', 'desc': '50/50 (50% PTFE + SS316)'},
                {'code': 'U', 'desc': 'UHMWPE'},
                {'code': 'P', 'desc': 'PEEK'},
                {'code': 'G', 'desc': 'Graphite Packing'},
                {'code': 'K', 'desc': 'Kel-F'}
            ],
            'bodySeal': [
                {'code': 'G', 'desc': 'Graphite Gasket'},
                {'code': 'F', 'desc': 'TFM Gasket'},
                {'code': 'Y', 'desc': 'CTFM Gasket'},
                {'code': 'T', 'desc': 'PTFE Gasket'},
                {'code': 'R', 'desc': 'RPTFE Gasket'},
                {'code': 'S', 'desc': '50/50 Gasket'},
                {'code': 'U', 'desc': 'UHMWPE Gasket'},
                {'code': 'P', 'desc': 'PEEK Gasket'},
                {'code': 'K', 'desc': 'Kel-F Gasket'}
            ],
            'oRings': [
                {'code': 'V', 'desc': 'VITON O-Rings'},
                {'code': 'E', 'desc': 'EPDM O-Rings'},
                {'code': 'T', 'desc': 'PTFE O-Rings'},
                {'code': 'B', 'desc': 'BUNA O-Rings'},
                {'code': 'N', 'desc': 'NONE'}
            ],
            'size': [
                {'code': '15', 'desc': '1/2 inch (15mm)'},
                {'code': '20', 'desc': '3/4 inch (20mm)'},
                {'code': '25', 'desc': '1 inch (25mm)'},
                {'code': '32', 'desc': '1-1/4 inch (32mm)'},
                {'code': '40', 'desc': '1-1/2 inch (40mm)'},
                {'code': '50', 'desc': '2 inch (50mm)'},
                {'code': '65', 'desc': '2-1/2 inch (65mm)'},
                {'code': '80', 'desc': '3 inch (80mm)'},
                {'code': '100', 'desc': '4 inch (100mm)'},
                {'code': '125', 'desc': '5 inch (125mm)'},
                {'code': '150', 'desc': '6 inch (150mm)'},
                {'code': '200', 'desc': '8 inch (200mm)'},
                {'code': '250', 'desc': '10 inch (250mm)'},
                {'code': '300', 'desc': '12 inch (300mm)'}
            ],
            'operator': [
                {'code': 'L', 'desc': 'Lever Locking'},
                {'code': 'O', 'desc': 'Oval Locking'},
                {'code': 'G', 'desc': 'Gear Operator'},
                {'code': 'S', 'desc': 'Deadman Operator'},
                {'code': 'A', 'desc': 'Actuator'},
                {'code': 'N', 'desc': 'Bare Stem'},
                {'code': 'X', 'desc': 'Special'}
            ]
        }
        
        self.setup_ui()
        
    def setup_ui(self):
        # Main container
        main_frame = tk.Frame(self.root, bg='#f8fafc')
        main_frame.pack(fill='both', expand=True, padx=20, pady=20)
        
        # Header
        header_frame = tk.Frame(main_frame, bg='white', relief='raised', bd=1)
        header_frame.pack(fill='x', pady=(0, 20))
        
        title_label = tk.Label(header_frame, text="Flo-Tite Valve Configurator", 
                              font=('Arial', 18, 'bold'), bg='white', fg='#1e293b')
        title_label.pack(pady=10)
        
        subtitle_label = tk.Label(header_frame, text="Configure your valve model number by selecting options below", 
                                 font=('Arial', 10), bg='white', fg='#64748b')
        subtitle_label.pack(pady=(0, 10))
        
        # Model number display
        model_frame = tk.Frame(header_frame, bg='#dbeafe', relief='raised', bd=2)
        model_frame.pack(fill='x', padx=10, pady=10)
        
        model_label = tk.Label(model_frame, text="Generated Model Number", 
                              font=('Arial', 10, 'bold'), bg='#dbeafe', fg='#1d4ed8')
        model_label.pack(anchor='w', padx=10, pady=(10, 0))
        
        self.model_number_var = tk.StringVar()
        self.model_number_label = tk.Label(model_frame, textvariable=self.model_number_var, 
                                          font=('Courier', 16, 'bold'), bg='#dbeafe', fg='#1e3a8a')
        self.model_number_label.pack(anchor='w', padx=10, pady=(0, 10))
        
        # Copy button
        copy_button = tk.Button(model_frame, text="Copy", command=self.copy_to_clipboard,
                               bg='#2563eb', fg='white', font=('Arial', 10, 'bold'),
                               relief='raised', bd=2)
        copy_button.pack(side='right', padx=10, pady=10)
        
        # Configuration options
        config_frame = tk.Frame(main_frame, bg='white', relief='raised', bd=1)
        config_frame.pack(fill='both', expand=True)
        
        # Create scrollable frame
        canvas = tk.Canvas(config_frame, bg='white')
        scrollbar = ttk.Scrollbar(config_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg='white')
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Configuration widgets
        self.create_config_widgets(scrollable_frame)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Footer
        footer_frame = tk.Frame(main_frame, bg='#f8fafc')
        footer_frame.pack(fill='x', pady=(20, 0))
        
        footer_label = tk.Label(footer_frame, text="Flo-Tite Model Number Configurator\nBased on Tech Bulletin Page 58-25", 
                               font=('Arial', 9), bg='#f8fafc', fg='#64748b')
        footer_label.pack()
        
        # Initialize model number
        self.update_model_number()
        
    def create_config_widgets(self, parent):
        # Model
        self.create_dropdown(parent, "1. Model Series", "model", 0)
        
        # Body Material
        self.create_dropdown(parent, "2. Body Material", "bodyMaterial", 1)
        
        # Seat Material
        self.create_dropdown(parent, "3. Seat Material", "seat", 2)
        
        # Stem Seal
        self.create_dropdown(parent, "4. Stem Seal Material", "stemSeal", 3)
        
        # Body Seal
        self.create_dropdown(parent, "5. Body Seal/Gasket Material", "bodySeal", 4)
        
        # O-Rings
        self.create_dropdown(parent, "6. O-Ring Material", "oRings", 5)
        
        # Operator
        self.create_dropdown(parent, "7. Operator Type", "operator", 6)
        
        # Size
        self.create_dropdown(parent, "8. Valve Size", "size", 7)
        
        # Special Features
        self.create_special_input(parent, "9. Special Features (Optional)", 8)
        
    def create_dropdown(self, parent, label_text, field, row):
        frame = tk.Frame(parent, bg='white')
        frame.pack(fill='x', padx=20, pady=10)
        
        label = tk.Label(frame, text=label_text, font=('Arial', 10, 'bold'), 
                        bg='white', fg='#1e293b')
        label.pack(anchor='w')
        
        var = tk.StringVar(value=self.config[field])
        combo = ttk.Combobox(frame, textvariable=var, state='readonly', width=80)
        combo.pack(fill='x', pady=(5, 0))
        
        # Populate options
        options_list = []
        for opt in self.options[field]:
            options_list.append(f"{opt['code']} - {opt['desc']}")
        combo['values'] = options_list
        
        # Set current value
        current_option = next((opt for opt in self.options[field] if opt['code'] == self.config[field]), None)
        if current_option:
            combo.set(f"{current_option['code']} - {current_option['desc']}")
        
        # Bind change event
        combo.bind('<<ComboboxSelected>>', lambda e, f=field: self.on_dropdown_change(f, combo.get()))
        
        # Store reference for updates
        setattr(self, f'{field}_combo', combo)
        
    def create_special_input(self, parent, label_text, row):
        frame = tk.Frame(parent, bg='white')
        frame.pack(fill='x', padx=20, pady=10)
        
        label = tk.Label(frame, text=label_text, font=('Arial', 10, 'bold'), 
                        bg='white', fg='#1e293b')
        label.pack(anchor='w')
        
        self.special_var = tk.StringVar(value=self.config['special'])
        special_entry = tk.Entry(frame, textvariable=self.special_var, width=80, font=('Arial', 10))
        special_entry.pack(fill='x', pady=(5, 0))
        special_entry.insert(0, "e.g., 3/8 GAUGE, V50, etc.")
        special_entry.configure(fg='gray')
        
        def on_focus_in(event):
            if special_entry.get() == "e.g., 3/8 GAUGE, V50, etc.":
                special_entry.delete(0, tk.END)
                special_entry.configure(fg='black')
        
        def on_focus_out(event):
            if not special_entry.get():
                special_entry.insert(0, "e.g., 3/8 GAUGE, V50, etc.")
                special_entry.configure(fg='gray')
        
        special_entry.bind('<FocusIn>', on_focus_in)
        special_entry.bind('<FocusOut>', on_focus_out)
        special_entry.bind('<KeyRelease>', lambda e: self.on_special_change())
        
        help_label = tk.Label(frame, text="Examples: Gauge taps, V-ports, stem extensions, etc.", 
                               font=('Arial', 8), bg='white', fg='#64748b')
        help_label.pack(anchor='w', pady=(2, 0))
        
    def on_dropdown_change(self, field, value):
        # Extract code from "CODE - Description" format
        code = value.split(' - ')[0]
        self.config[field] = code
        self.update_model_number()
        
    def on_special_change(self):
        value = self.special_var.get()
        if value != "e.g., 3/8 GAUGE, V50, etc.":
            self.config['special'] = value
        else:
            self.config['special'] = ''
        self.update_model_number()
        
    def generate_model_number(self):
        parts = [
            self.config['model'],
            self.config['bodyMaterial'],
            self.config['seat'],
            self.config['stemSeal'],
            self.config['bodySeal'],
            self.config['oRings'],
            self.config['operator'],
            self.config['size']
        ]
        
        model_number = '-'.join(parts)
        if self.config['special']:
            model_number += f"-{self.config['special']}"
        return model_number
        
    def update_model_number(self):
        model_number = self.generate_model_number()
        self.model_number_var.set(model_number)
        
    def copy_to_clipboard(self):
        model_number = self.generate_model_number()
        try:
            pyperclip.copy(model_number)
            messagebox.showinfo("Copied", f"Model number copied to clipboard:\n{model_number}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to copy to clipboard: {str(e)}")

def main():
    root = tk.Tk()
    app = FlotiteValveConfigurator(root)
    root.mainloop()

if __name__ == "__main__":
    main()
