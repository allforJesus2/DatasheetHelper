import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import os
import re

class ExcelMacroViewer:
    def __init__(self, master):
        self.master = master
        self.master.title("Excel Macros")
        self.master.geometry("400x500") # Adjusted size for potentially many buttons
        self.master.transient(master.master) # Make it transient to the main app window
        self.master.grab_set() # Make it modal

        self.macros = {} # Dictionary to store {macro_name: macro_content}

        self._load_macros()
        self._create_widgets()

    def _load_macros(self):
        # Assume 'excel_macros' folder is in the current working directory
        macros_dir = os.path.join(os.getcwd(), "excel_macros")
        self.macros.clear() # Clear existing macros

        if not os.path.isdir(macros_dir):
            print(f"Macro directory not found: {macros_dir}")
            # Optionally display this message in the GUI as well
            return

        print(f"Loading macros from: {macros_dir}")
        for filename in os.listdir(macros_dir):
            if filename.lower().endswith(".txt"):
                file_path = os.path.join(macros_dir, filename)
                try:
                    with open(file_path, 'r', encoding='utf-8') as f:
                        content = f.read()
                    
                    # Extract actual macro name from content
                    # Look for pattern like "Sub MacroName()" or "Function MacroName()"
                    match = re.search(r'(?:Sub|Function)\s+(\w+)\s*\(', content, re.IGNORECASE)
                    if match:
                        macro_name = match.group(1)
                    else:
                        # Fallback to filename if no match found
                        macro_name = os.path.splitext(filename)[0]
                        print(f"Warning: Could not extract macro name from {filename}, using filename instead.")
                    
                    if macro_name in self.macros:
                        print(f"Warning: Duplicate macro name '{macro_name}'. Overwriting.")
                    self.macros[macro_name] = content
                    print(f"  Loaded macro: {macro_name}")
                except Exception as e:
                    print(f"Error reading macro file {filename}: {e}")

    def _create_widgets(self):
        # Main frame
        main_frame = ttk.Frame(self.master, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # --- Scrollable Button Area ---
        canvas = tk.Canvas(main_frame)
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(
                scrollregion=canvas.bbox("all")
            )
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        # --- End Scrollable Area ---

        if not self.macros:
            ttk.Label(scrollable_frame, text="No macros found in 'excel_macros' folder.").pack(pady=20)
            return

        # Create buttons for each macro
        for name, content in sorted(self.macros.items()): # Sort alphabetically
            # Use lambda with default argument to capture correct content
            btn = ttk.Button(scrollable_frame, text=name,
                             command=lambda c=content, n=name: self._copy_to_clipboard(c, n))
            btn.pack(pady=3, padx=5, fill=tk.X) # Fill horizontally

        # Add a close button
        close_button = ttk.Button(main_frame, text="Close", command=self.master.destroy)
        # Place close button below the scrollbar
        close_button.pack(side=tk.BOTTOM, pady=(10, 0))


    def _copy_to_clipboard(self, content, name):
        try:
            self.master.clipboard_clear()
            self.master.clipboard_append(content)
            print(f"Macro '{name}' copied to clipboard.")
        except Exception as e:
            print(f"Error copying macro '{name}' to clipboard: {e}")

# Example usage (for testing purposes)
if __name__ == '__main__':
    # Create a dummy excel_macros folder and file for testing
    if not os.path.exists("excel_macros"):
        os.makedirs("excel_macros")
    with open("excel_macros/TestMacro1.txt", "w") as f:
        f.write("Sub TestMacro1()\n    MsgBox \"This is Test Macro 1\"\nEnd Sub")
    with open("excel_macros/Another_Macro.txt", "w") as f:
        f.write("Sub Another_Macro()\n    Range(\"A1\").Value = \"Hello\"\nEnd Sub")

    root = tk.Tk()
    root.withdraw() # Hide the main root window if only testing the dialog
    app = ExcelMacroViewer(tk.Toplevel(root))
    root.mainloop()

    # Clean up dummy folder/files
    # import shutil
    # if os.path.exists("excel_macros"):
    #     shutil.rmtree("excel_macros") 