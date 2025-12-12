import tkinter as tk
from tkinter import scrolledtext, messagebox, ttk
import difflib
from tkinter import font as tkFont


class TextComparisonApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Text Comparison Tool")
        self.root.geometry("1200x800")
        self.root.minsize(800, 600)
        
        # Configure style
        style = ttk.Style()
        style.theme_use('clam')
        
        # Create main container
        main_frame = ttk.Frame(root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)
        
        # Title
        title_label = ttk.Label(main_frame, text="Text Comparison Tool", 
                                font=('Arial', 16, 'bold'))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 10))
        
        # Left panel - Text 1
        left_frame = ttk.LabelFrame(main_frame, text="Text 1", padding="5")
        left_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 5))
        
        self.text1 = scrolledtext.ScrolledText(left_frame, wrap=tk.WORD, 
                                               width=40, height=20,
                                               font=('Consolas', 10))
        self.text1.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.text1.bind('<KeyRelease>', self.on_text_change)
        self.text1.bind('<Button-1>', self.on_text_change)
        self.text1.bind('<Control-v>', self.on_text_change)
        
        left_frame.columnconfigure(0, weight=1)
        left_frame.rowconfigure(0, weight=1)
        
        # Middle panel - Comparison Results
        middle_frame = ttk.LabelFrame(main_frame, text="Differences", padding="5")
        middle_frame.grid(row=1, column=1, sticky=(tk.W, tk.E, tk.N, tk.S), padx=5)
        
        self.diff_text = scrolledtext.ScrolledText(middle_frame, wrap=tk.WORD,
                                                   width=40, height=20,
                                                   font=('Consolas', 10),
                                                   state=tk.DISABLED)
        self.diff_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        middle_frame.columnconfigure(0, weight=1)
        middle_frame.rowconfigure(0, weight=1)
        
        # Right panel - Text 2
        right_frame = ttk.LabelFrame(main_frame, text="Text 2", padding="5")
        right_frame.grid(row=1, column=2, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(5, 0))
        
        self.text2 = scrolledtext.ScrolledText(right_frame, wrap=tk.WORD,
                                               width=40, height=20,
                                               font=('Consolas', 10))
        self.text2.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.text2.bind('<KeyRelease>', self.on_text_change)
        self.text2.bind('<Button-1>', self.on_text_change)
        self.text2.bind('<Control-v>', self.on_text_change)
        
        right_frame.columnconfigure(0, weight=1)
        right_frame.rowconfigure(0, weight=1)
        
        # Configure grid weights for resizing
        main_frame.columnconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.columnconfigure(2, weight=1)
        main_frame.rowconfigure(1, weight=1)
        
        # Button frame
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=2, column=0, columnspan=3, pady=(10, 0))
        
        # Compare button
        compare_btn = ttk.Button(button_frame, text="Compare Texts", 
                                command=self.compare_texts)
        compare_btn.pack(side=tk.LEFT, padx=5)
        
        # Clear button
        clear_btn = ttk.Button(button_frame, text="Clear All", 
                              command=self.clear_all)
        clear_btn.pack(side=tk.LEFT, padx=5)
        
        # Copy differences button
        copy_btn = ttk.Button(button_frame, text="Copy Differences", 
                            command=self.copy_differences)
        copy_btn.pack(side=tk.LEFT, padx=5)
        
        # Statistics label
        self.stats_label = ttk.Label(button_frame, text="", foreground="gray")
        self.stats_label.pack(side=tk.LEFT, padx=20)
        
        # Configure text tags for highlighting
        self.diff_text.tag_config('added', background='#90EE90', foreground='black')
        self.diff_text.tag_config('deleted', background='#FFB6C1', foreground='black')
        self.diff_text.tag_config('changed', background='#FFD700', foreground='black')
        self.diff_text.tag_config('equal', foreground='gray')
        
        # Enable text widgets for copy/paste
        self.text1.bind('<Control-c>', self.copy_text1)
        self.text2.bind('<Control-c>', self.copy_text2)
        self.diff_text.bind('<Control-c>', self.copy_diff)
        
        # Initial comparison
        self.compare_texts()
    
    def on_text_change(self, event=None):
        # Auto-compare after a short delay (debounce)
        self.root.after(500, self.compare_texts)
    
    def compare_texts(self):
        """Compare the two texts and display differences"""
        text1_content = self.text1.get("1.0", tk.END).strip()
        text2_content = self.text2.get("1.0", tk.END).strip()
        
        # Enable diff text widget for editing
        self.diff_text.config(state=tk.NORMAL)
        self.diff_text.delete("1.0", tk.END)
        
        if not text1_content and not text2_content:
            self.diff_text.insert("1.0", "Paste or type text in both panels to compare.")
            self.diff_text.config(state=tk.DISABLED)
            self.stats_label.config(text="")
            return
        
        if not text1_content:
            self.diff_text.insert("1.0", "Text 1 is empty.\n\n", "deleted")
            self.diff_text.insert(tk.END, text2_content, "added")
            self.diff_text.config(state=tk.DISABLED)
            self.stats_label.config(text="Text 1 is empty")
            return
        
        if not text2_content:
            self.diff_text.insert("1.0", text1_content, "deleted")
            self.diff_text.insert(tk.END, "\n\nText 2 is empty.", "added")
            self.diff_text.config(state=tk.DISABLED)
            self.stats_label.config(text="Text 2 is empty")
            return
        
        # Split texts into lines for comparison
        lines1 = text1_content.splitlines(keepends=True)
        lines2 = text2_content.splitlines(keepends=True)
        
        # Use difflib to find differences
        diff = difflib.unified_diff(lines1, lines2, 
                                   fromfile='Text 1', 
                                   tofile='Text 2',
                                   lineterm='', n=0)
        
        diff_list = list(diff)
        
        if len(diff_list) == 0:
            self.diff_text.insert("1.0", "✓ Texts are identical!", "equal")
            self.diff_text.config(state=tk.DISABLED)
            self.stats_label.config(text="✓ Texts are identical")
            return
        
        # Process and display differences
        self.display_differences(lines1, lines2)
        
        # Calculate statistics
        self.calculate_statistics(lines1, lines2)
    
    def highlight_character_differences(self, old_line, new_line, show_old=True):
        """Perform character-level comparison and return formatted text with tags.
        Returns a list of tuples: (text, tag) where tag can be None for default styling.
        
        Args:
            old_line: The original line to compare
            new_line: The new line to compare
            show_old: If True, return segments from old_line (for deleted lines);
                     If False, return segments from new_line (for added lines)
        """
        if not old_line and not new_line:
            return [("", None)]
        
        if old_line == new_line:
            line_to_show = old_line if show_old else new_line
            return [(line_to_show, 'equal')]
        
        # Remove line endings for comparison, we'll add them back
        old_stripped = old_line.rstrip('\n\r')
        new_stripped = new_line.rstrip('\n\r')
        
        # Determine which line and ending to use based on show_old
        if show_old:
            line_stripped = old_stripped
            line_ending = old_line[len(old_stripped):] if len(old_line) > len(old_stripped) else ""
        else:
            line_stripped = new_stripped
            line_ending = new_line[len(new_stripped):] if len(new_line) > len(new_stripped) else ""
        
        # Use SequenceMatcher for character-level comparison
        matcher = difflib.SequenceMatcher(None, old_stripped, new_stripped)
        segments = []
        
        for tag, i1, i2, j1, j2 in matcher.get_opcodes():
            if tag == 'equal':
                # Show equal parts from the line we're displaying
                if show_old:
                    segments.append((old_stripped[i1:i2], 'equal'))
                else:
                    segments.append((new_stripped[j1:j2], 'equal'))
            elif tag == 'delete':
                # Only include if showing old line
                if show_old:
                    segments.append((old_stripped[i1:i2], 'deleted'))
            elif tag == 'insert':
                # Only include if showing new line
                if not show_old:
                    segments.append((new_stripped[j1:j2], 'added'))
            elif tag == 'replace':
                # Show replaced parts from the appropriate line
                if show_old:
                    segments.append((old_stripped[i1:i2], 'deleted'))
                else:
                    segments.append((new_stripped[j1:j2], 'added'))
        
        # Add line ending if present
        if line_ending:
            segments.append((line_ending, 'equal'))
        
        return segments
    
    def display_differences(self, lines1, lines2):
        """Display differences with color coding"""
        # Use SequenceMatcher for more detailed comparison
        matcher = difflib.SequenceMatcher(None, lines1, lines2)
        
        for tag, i1, i2, j1, j2 in matcher.get_opcodes():
            if tag == 'equal':
                # Show equal lines in gray
                for line in lines1[i1:i2]:
                    self.diff_text.insert(tk.END, line, 'equal')
            elif tag == 'delete':
                # Show deleted lines in red
                for line in lines1[i1:i2]:
                    self.diff_text.insert(tk.END, f"- {line}", 'deleted')
            elif tag == 'insert':
                # Show inserted lines in green
                for line in lines2[j1:j2]:
                    self.diff_text.insert(tk.END, f"+ {line}", 'added')
            elif tag == 'replace':
                # Show replaced lines with character-level differences
                old_lines_list = lines1[i1:i2]
                new_lines_list = lines2[j1:j2]
                
                # Show deleted lines (show old line with character-level differences)
                for idx, old_line in enumerate(old_lines_list):
                    self.diff_text.insert(tk.END, "- ")
                    # Compare old line with corresponding new line if available, otherwise with empty
                    if idx < len(new_lines_list):
                        new_line = new_lines_list[idx]
                    else:
                        new_line = ""
                    
                    segments = self.highlight_character_differences(old_line, new_line, show_old=True)
                    for text, seg_tag in segments:
                        if text:
                            self.diff_text.insert(tk.END, text, seg_tag if seg_tag else None)
                
                # Show added lines (show new line with character-level differences)
                for idx, new_line in enumerate(new_lines_list):
                    self.diff_text.insert(tk.END, "+ ")
                    # Compare new line with corresponding old line if available
                    if idx < len(old_lines_list):
                        old_line = old_lines_list[idx]
                    else:
                        old_line = ""
                    
                    segments = self.highlight_character_differences(old_line, new_line, show_old=False)
                    for text, seg_tag in segments:
                        if text:
                            self.diff_text.insert(tk.END, text, seg_tag if seg_tag else None)
        
        self.diff_text.config(state=tk.DISABLED)
    
    def calculate_statistics(self, lines1, lines2):
        """Calculate and display comparison statistics"""
        matcher = difflib.SequenceMatcher(None, lines1, lines2)
        similarity = matcher.ratio() * 100
        
        total_lines1 = len(lines1)
        total_lines2 = len(lines2)
        
        added = sum(1 for tag, _, _, _, _ in matcher.get_opcodes() if tag == 'insert')
        deleted = sum(1 for tag, _, _, _, _ in matcher.get_opcodes() if tag == 'delete')
        changed = sum(1 for tag, _, _, _, _ in matcher.get_opcodes() if tag == 'replace')
        
        stats = (f"Similarity: {similarity:.1f}% | "
                f"Text 1: {total_lines1} lines | "
                f"Text 2: {total_lines2} lines | "
                f"Added: {added} | Deleted: {deleted} | Changed: {changed}")
        
        self.stats_label.config(text=stats)
    
    def clear_all(self):
        """Clear all text areas"""
        self.text1.delete("1.0", tk.END)
        self.text2.delete("1.0", tk.END)
        self.diff_text.config(state=tk.NORMAL)
        self.diff_text.delete("1.0", tk.END)
        self.diff_text.config(state=tk.DISABLED)
        self.stats_label.config(text="")
    
    def copy_differences(self):
        """Copy the differences text to clipboard"""
        diff_content = self.diff_text.get("1.0", tk.END)
        if diff_content.strip():
            self.root.clipboard_clear()
            self.root.clipboard_append(diff_content)
            messagebox.showinfo("Copied", "Differences copied to clipboard!")
        else:
            messagebox.showwarning("Empty", "No differences to copy.")
    
    def copy_text1(self, event=None):
        """Copy selected text from text1"""
        try:
            text = self.text1.selection_get()
            self.root.clipboard_clear()
            self.root.clipboard_append(text)
        except tk.TclError:
            pass
        return "break"
    
    def copy_text2(self, event=None):
        """Copy selected text from text2"""
        try:
            text = self.text2.selection_get()
            self.root.clipboard_clear()
            self.root.clipboard_append(text)
        except tk.TclError:
            pass
        return "break"
    
    def copy_diff(self, event=None):
        """Copy selected text from diff area"""
        try:
            text = self.diff_text.selection_get()
            self.root.clipboard_clear()
            self.root.clipboard_append(text)
        except tk.TclError:
            pass
        return "break"


def main():
    root = tk.Tk()
    app = TextComparisonApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()

