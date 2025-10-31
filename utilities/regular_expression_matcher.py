import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
import re

class RegexMatcherGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Regular Expression Matcher")
        self.root.geometry("800x600")
        self.root.minsize(600, 400)
        
        # Configure grid weights for resizing
        self.root.grid_rowconfigure(2, weight=1)
        self.root.grid_rowconfigure(5, weight=1)
        self.root.grid_columnconfigure(0, weight=1)
        
        self.setup_ui()
    
    def setup_ui(self):
        # Title
        title_label = tk.Label(self.root, text="Regular Expression Matcher", 
                              font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, pady=10, sticky="ew")
        
        # Input header with label and Paste button
        input_header = tk.Frame(self.root)
        input_header.grid(row=1, column=0, sticky="ew", padx=10, pady=(0, 5))
        input_header.grid_columnconfigure(0, weight=1)

        input_label = tk.Label(input_header, text="Input Text:", font=("Arial", 12, "bold"))
        input_label.grid(row=0, column=0, sticky="w")

        paste_button = tk.Button(input_header, text="Paste", command=self.paste_into_input, font=("Arial", 10))
        paste_button.grid(row=0, column=1, sticky="e", padx=(10, 0))

        # Input text area
        self.input_text = scrolledtext.ScrolledText(
            self.root, 
            height=8, 
            wrap=tk.WORD,
            font=("Consolas", 10)
        )
        self.input_text.grid(row=2, column=0, sticky="nsew", padx=10, pady=(0, 10))
        
        # Regex input section
        regex_frame = tk.Frame(self.root)
        regex_frame.grid(row=3, column=0, sticky="ew", padx=10, pady=5)
        regex_frame.grid_columnconfigure(1, weight=1)
        
        regex_label = tk.Label(regex_frame, text="Regular Expression:", font=("Arial", 12, "bold"))
        regex_label.grid(row=0, column=0, sticky="w", padx=(0, 10))
        
        self.regex_entry = tk.Entry(regex_frame, font=("Consolas", 11), width=50)
        self.regex_entry.grid(row=0, column=1, sticky="ew", padx=(0, 10))
        self.regex_entry.bind('<Return>', lambda e: self.find_matches())
        
        # Match button
        self.match_button = tk.Button(
            regex_frame, 
            text="Find Matches", 
            command=self.find_matches,
            font=("Arial", 11, "bold"),
            bg="#4CAF50",
            fg="white",
            padx=20
        )
        self.match_button.grid(row=0, column=2, padx=(0, 10))
        
        # Clear button
        clear_button = tk.Button(
            regex_frame, 
            text="Clear All", 
            command=self.clear_all,
            font=("Arial", 11),
            bg="#f44336",
            fg="white",
            padx=20
        )
        clear_button.grid(row=0, column=3)
        
        # Output text area
        output_label = tk.Label(self.root, text="Matches (one per line):", font=("Arial", 12, "bold"))
        output_label.grid(row=4, column=0, sticky="w", padx=10, pady=(0, 5))
        
        self.output_text = scrolledtext.ScrolledText(
            self.root, 
            height=8, 
            wrap=tk.WORD,
            font=("Consolas", 10),
            state=tk.DISABLED
        )
        self.output_text.grid(row=5, column=0, sticky="nsew", padx=10, pady=(0, 10))

        # Copy button under output
        copy_button = tk.Button(self.root, text="Copy Matches", command=self.copy_output, font=("Arial", 10))
        copy_button.grid(row=6, column=0, sticky="e", padx=10, pady=(0, 5))
        
        # Status bar
        self.status_var = tk.StringVar()
        self.status_var.set("Ready - Enter text and regex pattern")
        status_bar = tk.Label(self.root, textvariable=self.status_var, 
                             relief=tk.SUNKEN, anchor=tk.W)
        status_bar.grid(row=7, column=0, sticky="ew", padx=10, pady=(0, 5))
        
        # Add some example regex patterns as tooltips
        self.add_tooltips()
    
    def add_tooltips(self):
        """Add helpful tooltips with common regex patterns"""
        tooltip_text = """Common Regex Patterns:
• \\d+ - One or more digits
• \\w+ - One or more word characters
• [a-zA-Z]+ - One or more letters
• \\b\\w+@\\w+\\.\\w+\\b - Email addresses
• \\b\\d{3}-\\d{3}-\\d{4}\\b - Phone numbers (xxx-xxx-xxxx)
• \\$\\d+\\.\\d{2} - Currency ($xx.xx)
• \\b[A-Z][a-z]+\\b - Capitalized words"""
        
        self.regex_entry.bind('<FocusIn>', lambda e: self.show_tooltip(tooltip_text))
        self.regex_entry.bind('<FocusOut>', lambda e: self.hide_tooltip())
    
    def show_tooltip(self, text):
        """Show tooltip with regex examples"""
        self.tooltip = tk.Toplevel(self.root)
        self.tooltip.wm_overrideredirect(True)
        self.tooltip.wm_geometry("+%d+%d" % (self.root.winfo_rootx() + 50, 
                                           self.root.winfo_rooty() + 100))
        
        tooltip_label = tk.Label(self.tooltip, text=text, 
                               justify=tk.LEFT, font=("Consolas", 9),
                               bg="lightyellow", relief=tk.SOLID, borderwidth=1)
        tooltip_label.pack()
    
    def hide_tooltip(self):
        """Hide tooltip"""
        if hasattr(self, 'tooltip'):
            self.tooltip.destroy()
    
    def find_matches(self):
        """Find and display regex matches"""
        try:
            # Get input text and regex pattern
            input_text = self.input_text.get("1.0", tk.END).strip()
            regex_pattern = self.regex_entry.get().strip()
            
            if not input_text:
                self.status_var.set("Error: Please enter some text to search")
                return
            
            if not regex_pattern:
                self.status_var.set("Error: Please enter a regular expression pattern")
                return
            
            # Compile regex pattern
            try:
                pattern = re.compile(regex_pattern)
            except re.error as e:
                self.status_var.set(f"Regex Error: {str(e)}")
                messagebox.showerror("Regex Error", f"Invalid regular expression:\n{str(e)}")
                return
            
            # Find all matches
            matches = pattern.findall(input_text)
            
            # Clear and populate output
            self.output_text.config(state=tk.NORMAL)
            self.output_text.delete("1.0", tk.END)
            
            if matches:
                # Join matches with newlines
                output_text = "\n".join(str(match) for match in matches)
                self.output_text.insert("1.0", output_text)
                self.status_var.set(f"Found {len(matches)} match(es)")
            else:
                self.output_text.insert("1.0", "No matches found")
                self.status_var.set("No matches found")
            
            self.output_text.config(state=tk.DISABLED)
            
        except Exception as e:
            self.status_var.set(f"Error: {str(e)}")
            messagebox.showerror("Error", f"An error occurred:\n{str(e)}")
    
    def paste_into_input(self):
        """Paste clipboard content into the input text at the cursor position"""
        try:
            clipboard_text = self.root.clipboard_get()
        except Exception:
            self.status_var.set("Clipboard is empty or unavailable")
            return
        self.input_text.insert(tk.INSERT, clipboard_text)
        self.status_var.set("Pasted clipboard into input")

    def copy_output(self):
        """Copy matches output to the clipboard"""
        # Temporarily enable to read content
        was_disabled = str(self.output_text.cget("state")) == str(tk.DISABLED)
        if was_disabled:
            self.output_text.config(state=tk.NORMAL)
        text = self.output_text.get("1.0", "end-1c").strip()
        if was_disabled:
            self.output_text.config(state=tk.DISABLED)
        if not text:
            self.status_var.set("Nothing to copy")
            return
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        self.status_var.set("Copied matches to clipboard")

    def clear_all(self):
        """Clear all text areas and input fields"""
        self.input_text.delete("1.0", tk.END)
        self.regex_entry.delete(0, tk.END)
        self.output_text.config(state=tk.NORMAL)
        self.output_text.delete("1.0", tk.END)
        self.output_text.config(state=tk.DISABLED)
        self.status_var.set("Cleared - Ready for new input")

def main():
    root = tk.Tk()
    app = RegexMatcherGUI(root)
    
    # Center the window on screen
    root.update_idletasks()
    x = (root.winfo_screenwidth() // 2) - (root.winfo_width() // 2)
    y = (root.winfo_screenheight() // 2) - (root.winfo_height() // 2)
    root.geometry(f"+{x}+{y}")
    
    root.mainloop()

if __name__ == "__main__":
    main()
