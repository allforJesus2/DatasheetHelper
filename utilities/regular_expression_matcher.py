import tkinter as tk
from tkinter import scrolledtext, messagebox, filedialog
import re
import os
from io import BytesIO
os.environ["KMP_DUPLICATE_LIB_OK"] = "TRUE"
# Import pdfplumber for PDF text extraction
try:
    import pdfplumber
except ImportError:
    pdfplumber = None

# Import libraries for OCR fallback
try:
    import fitz  # PyMuPDF
    fitz_available = True
except ImportError:
    fitz_available = False

try:
    from PIL import Image
    import numpy as np
    pil_available = True
    numpy_available = True
except ImportError:
    pil_available = False
    numpy_available = False

try:
    import easyocr
    easyocr_available = True
    # Initialize EasyOCR reader (lazy loading)
    easyocr_reader = None
except ImportError:
    easyocr_available = False
    easyocr_reader = None

class RegexMatcherGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Regular Expression Matcher")
        self.root.geometry("800x600")
        self.root.minsize(600, 400)
        
        # Configure grid weights for resizing
        self.root.grid_rowconfigure(3, weight=1)
        self.root.grid_rowconfigure(6, weight=1)
        self.root.grid_columnconfigure(0, weight=1)
        
        self.setup_ui()
    
    def setup_ui(self):
        # PDF extraction section at the top
        pdf_frame = tk.Frame(self.root, relief=tk.RAISED, bd=1)
        pdf_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=10)
        pdf_frame.grid_columnconfigure(1, weight=1)
        
        pdf_label = tk.Label(pdf_frame, text="PDF Path:", font=("Arial", 10, "bold"))
        pdf_label.grid(row=0, column=0, sticky="w", padx=(5, 10))
        
        self.pdf_path_entry = tk.Entry(pdf_frame, font=("Consolas", 10))
        self.pdf_path_entry.grid(row=0, column=1, sticky="ew", padx=(0, 5))
        self.pdf_path_entry.bind('<Return>', lambda e: self.extract_pdf_text())
        
        browse_pdf_button = tk.Button(
            pdf_frame, 
            text="Browse", 
            command=self.browse_pdf_file,
            font=("Arial", 10)
        )
        browse_pdf_button.grid(row=0, column=2, padx=(0, 5))
        
        extract_button = tk.Button(
            pdf_frame, 
            text="Extract", 
            command=self.extract_pdf_text,
            font=("Arial", 10, "bold"),
            bg="#2196F3",
            fg="white",
            padx=15
        )
        extract_button.grid(row=0, column=3, padx=(0, 5))
        
        # Title
        title_label = tk.Label(self.root, text="Regular Expression Matcher", 
                              font=("Arial", 16, "bold"))
        title_label.grid(row=1, column=0, pady=10, sticky="ew")
        
        # Input header with label and Paste button
        input_header = tk.Frame(self.root)
        input_header.grid(row=2, column=0, sticky="ew", padx=10, pady=(0, 5))
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
        self.input_text.grid(row=3, column=0, sticky="nsew", padx=10, pady=(0, 10))
        
        # Regex input section
        regex_frame = tk.Frame(self.root)
        regex_frame.grid(row=4, column=0, sticky="ew", padx=10, pady=5)
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
        output_label.grid(row=5, column=0, sticky="w", padx=10, pady=(0, 5))
        
        self.output_text = scrolledtext.ScrolledText(
            self.root, 
            height=8, 
            wrap=tk.WORD,
            font=("Consolas", 10),
            state=tk.DISABLED
        )
        self.output_text.grid(row=6, column=0, sticky="nsew", padx=10, pady=(0, 10))

        # Copy button under output
        copy_button = tk.Button(self.root, text="Copy Matches", command=self.copy_output, font=("Arial", 10))
        copy_button.grid(row=7, column=0, sticky="e", padx=10, pady=(0, 5))
        
        # Status bar
        self.status_var = tk.StringVar()
        self.status_var.set("Ready - Enter text and regex pattern")
        status_bar = tk.Label(self.root, textvariable=self.status_var, 
                             relief=tk.SUNKEN, anchor=tk.W)
        status_bar.grid(row=8, column=0, sticky="ew", padx=10, pady=(0, 5))
        
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

    def browse_pdf_file(self):
        """Open file dialog to browse for PDF file"""
        file_path = filedialog.askopenfilename(
            title="Select PDF File",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if file_path:
            self.pdf_path_entry.delete(0, tk.END)
            self.pdf_path_entry.insert(0, file_path)
            self.status_var.set(f"Selected PDF: {os.path.basename(file_path)}")
    
    def extract_pdf_text_with_ocr(self, pdf_path):
        """Extract text from PDF using OCR when no text is found"""
        global easyocr_reader
        
        if not fitz_available:
            raise ImportError("PyMuPDF (fitz) library is not installed. Install with: pip install PyMuPDF")
        
        if not pil_available or not numpy_available:
            raise ImportError("Pillow and numpy libraries are required. Install with: pip install Pillow numpy")
        
        if not easyocr_available:
            raise ImportError("easyocr library is not installed. Install with: pip install easyocr")
        
        # Initialize EasyOCR reader if not already done
        if easyocr_reader is None:
            self.status_var.set("Initializing OCR engine (first time only, this may take a moment)...")
            self.root.update()
            easyocr_reader = easyocr.Reader(['en'], gpu=True)
        
        self.status_var.set("Converting PDF to images...")
        self.root.update()
        
        # Convert PDF to images using fitz
        try:
            pdf_document = fitz.open(pdf_path)
            total_pages = len(pdf_document)
        except Exception as e:
            raise Exception(f"Failed to open PDF: {str(e)}")
        
        extracted_text = ""
        
        try:
            for page_num in range(total_pages):
                self.status_var.set(f"Performing OCR on page {page_num + 1}/{total_pages}...")
                self.root.update()
                
                # Get the page
                page = pdf_document[page_num]
                
                # Render page to a pixmap (image)
                pixmap = page.get_pixmap()
                
                # Convert pixmap to numpy array for EasyOCR
                img_data = pixmap.tobytes("ppm")
                image = Image.open(BytesIO(img_data))
                img_array = np.array(image)
                
                # Perform OCR on the numpy array
                results = easyocr_reader.readtext(img_array)
                
                # Combine all text from OCR results
                page_text = "\n".join([result[1] for result in results])
                if page_text:
                    extracted_text += page_text + "\n\n"
        finally:
            pdf_document.close()
        
        return extracted_text.strip()
    
    def extract_pdf_text(self):
        """Extract text from PDF and populate the input text area"""
        pdf_path = self.pdf_path_entry.get().strip()
        
        if not pdf_path:
            self.status_var.set("Error: Please enter or browse for a PDF file path")
            messagebox.showwarning("No PDF Path", "Please enter or browse for a PDF file path.")
            return
        
        if not os.path.exists(pdf_path):
            self.status_var.set(f"Error: PDF file not found: {pdf_path}")
            messagebox.showerror("File Not Found", f"The file does not exist:\n{pdf_path}")
            return
        
        if not pdf_path.lower().endswith('.pdf'):
            self.status_var.set("Error: Selected file is not a PDF")
            messagebox.showerror("Invalid File", "Please select a PDF file (.pdf)")
            return
        
        # Check if pdfplumber is available
        if pdfplumber is None:
            self.status_var.set("Error: pdfplumber library not available")
            messagebox.showerror(
                "PDF Parser Not Available",
                "pdfplumber library is not installed.\n\n"
                "Please install it with:\n"
                "  pip install pdfplumber"
            )
            return
        
        try:
            self.status_var.set("Extracting text from PDF...")
            self.root.update()  # Update UI to show status
            
            extracted_text = ""
            with pdfplumber.open(pdf_path) as pdf:
                for page in pdf.pages:
                    page_text = page.extract_text()
                    if page_text:
                        extracted_text += page_text + "\n\n"
            
            if extracted_text.strip():
                # Clear existing input and insert extracted text
                self.input_text.delete("1.0", tk.END)
                self.input_text.insert("1.0", extracted_text.strip())
                self.status_var.set(f"Successfully extracted text from PDF ({len(extracted_text)} characters)")
            else:
                # No text found, try OCR fallback
                self.status_var.set("No text found, attempting OCR...")
                self.root.update()
                
                try:
                    # Ask user if they want to use OCR
                    use_ocr = messagebox.askyesno(
                        "No Text Found",
                        "No text could be extracted from the PDF.\n\n"
                        "This may be an image-based PDF. Would you like to use OCR (Optical Character Recognition) to extract text?\n\n"
                        "Note: OCR may take longer and requires additional libraries."
                    )
                    
                    if use_ocr:
                        extracted_text = self.extract_pdf_text_with_ocr(pdf_path)
                        
                        if extracted_text:
                            # Clear existing input and insert extracted text
                            self.input_text.delete("1.0", tk.END)
                            self.input_text.insert("1.0", extracted_text)
                            self.status_var.set(f"Successfully extracted text using OCR ({len(extracted_text)} characters)")
                        else:
                            self.status_var.set("OCR completed but no text was found")
                            messagebox.showwarning(
                                "OCR Complete",
                                "OCR processing completed but no text was extracted from the PDF images."
                            )
                    else:
                        self.status_var.set("Text extraction cancelled by user")
                except ImportError as e:
                    self.status_var.set(f"OCR Error: {str(e)}")
                    messagebox.showerror(
                        "OCR Libraries Not Available",
                        f"{str(e)}\n\n"
                        "Please install required libraries:\n"
                        "  pip install PyMuPDF Pillow numpy easyocr"
                    )
                except Exception as e:
                    error_msg = f"Error during OCR: {str(e)}"
                    self.status_var.set(error_msg)
                    messagebox.showerror("OCR Error", error_msg)
        
        except Exception as e:
            error_msg = f"Error extracting PDF text: {str(e)}"
            self.status_var.set(error_msg)
            messagebox.showerror("PDF Extraction Error", error_msg)
    
    def clear_all(self):
        """Clear all text areas and input fields"""
        self.input_text.delete("1.0", tk.END)
        self.regex_entry.delete(0, tk.END)
        self.pdf_path_entry.delete(0, tk.END)
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
