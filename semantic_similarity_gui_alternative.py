import tkinter as tk
from tkinter import ttk, messagebox
import numpy as np
from typing import Tuple
import threading
import re
from collections import Counter

class SimpleSemanticSimilarityGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Simple Semantic Similarity Calculator")
        self.root.geometry("600x450")
        self.root.resizable(True, True)
        
        # Create the main frame
        self.main_frame = ttk.Frame(root, padding="20")
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Create widgets
        self.create_widgets()
    
    def create_widgets(self):
        # Title
        title_label = ttk.Label(self.main_frame, text="Simple Semantic Similarity Calculator", 
                               font=("Arial", 16, "bold"))
        title_label.pack(pady=(0, 20))
        
        # Description
        desc_label = ttk.Label(self.main_frame, 
                              text="Calculate similarity using multiple methods: Jaccard, Cosine, and Edit Distance",
                              font=("Arial", 10), wraplength=500)
        desc_label.pack(pady=(0, 20))
        
        # Input frame
        input_frame = ttk.LabelFrame(self.main_frame, text="Input Words", padding="10")
        input_frame.pack(fill=tk.X, pady=(0, 20))
        
        # Word 1
        word1_frame = ttk.Frame(input_frame)
        word1_frame.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(word1_frame, text="Word 1:").pack(side=tk.LEFT)
        self.word1_entry = ttk.Entry(word1_frame, width=40)
        self.word1_entry.pack(side=tk.LEFT, padx=(10, 0), fill=tk.X, expand=True)
        
        # Word 2
        word2_frame = ttk.Frame(input_frame)
        word2_frame.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(word2_frame, text="Word 2:").pack(side=tk.LEFT)
        self.word2_entry = ttk.Entry(word2_frame, width=40)
        self.word2_entry.pack(side=tk.LEFT, padx=(10, 0), fill=tk.X, expand=True)
        
        # Calculate button
        self.calculate_button = ttk.Button(input_frame, text="Calculate Similarity", 
                                          command=self.calculate_similarity)
        self.calculate_button.pack(pady=(10, 0))
        
        # Results frame
        results_frame = ttk.LabelFrame(self.main_frame, text="Results", padding="10")
        results_frame.pack(fill=tk.BOTH, expand=True)
        
        # Create a notebook for different similarity methods
        self.notebook = ttk.Notebook(results_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        
        # Jaccard similarity tab
        self.jaccard_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.jaccard_frame, text="Jaccard")
        self.create_jaccard_tab()
        
        # Cosine similarity tab
        self.cosine_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.cosine_frame, text="Cosine")
        self.create_cosine_tab()
        
        # Edit distance tab
        self.edit_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.edit_frame, text="Edit Distance")
        self.create_edit_tab()
        
        # Examples frame
        examples_frame = ttk.LabelFrame(self.main_frame, text="Quick Examples", padding="10")
        examples_frame.pack(fill=tk.X, pady=(20, 0))
        
        # Example buttons
        examples = [
            ("cat", "dog"),
            ("happy", "sad"),
            ("computer", "laptop"),
            ("car", "bicycle"),
            ("apple", "orange"),
            ("hello", "world")
        ]
        
        for i, (word1, word2) in enumerate(examples):
            btn = ttk.Button(examples_frame, text=f"{word1} vs {word2}", 
                           command=lambda w1=word1, w2=word2: self.load_example(w1, w2))
            btn.grid(row=i//3, column=i%3, padx=5, pady=2, sticky="ew")
        
        # Configure grid weights
        for i in range(3):
            examples_frame.columnconfigure(i, weight=1)
    
    def create_jaccard_tab(self):
        """Create the Jaccard similarity tab"""
        # Score display
        score_frame = ttk.Frame(self.jaccard_frame)
        score_frame.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(score_frame, text="Jaccard Similarity:").pack(side=tk.LEFT)
        self.jaccard_score_label = ttk.Label(score_frame, text="--", font=("Arial", 12, "bold"))
        self.jaccard_score_label.pack(side=tk.LEFT, padx=(10, 0))
        
        # Progress bar
        self.jaccard_progress_var = tk.DoubleVar()
        self.jaccard_progress_bar = ttk.Progressbar(self.jaccard_frame, variable=self.jaccard_progress_var, 
                                                   maximum=1.0, length=300)
        self.jaccard_progress_bar.pack(pady=(0, 10))
        
        # Interpretation
        self.jaccard_interpretation_label = ttk.Label(self.jaccard_frame, text="", 
                                                     font=("Arial", 10), wraplength=500)
        self.jaccard_interpretation_label.pack(pady=(0, 10))
        
        # Explanation
        explanation = ("Jaccard Similarity measures the overlap between character sets of two words.\n"
                      "Formula: |A ∩ B| / |A ∪ B| where A and B are character sets.")
        ttk.Label(self.jaccard_frame, text=explanation, font=("Arial", 9), 
                wraplength=500, foreground="gray").pack()
    
    def create_cosine_tab(self):
        """Create the Cosine similarity tab"""
        # Score display
        score_frame = ttk.Frame(self.cosine_frame)
        score_frame.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(score_frame, text="Cosine Similarity:").pack(side=tk.LEFT)
        self.cosine_score_label = ttk.Label(score_frame, text="--", font=("Arial", 12, "bold"))
        self.cosine_score_label.pack(side=tk.LEFT, padx=(10, 0))
        
        # Progress bar
        self.cosine_progress_var = tk.DoubleVar()
        self.cosine_progress_bar = ttk.Progressbar(self.cosine_frame, variable=self.cosine_progress_var, 
                                                  maximum=1.0, length=300)
        self.cosine_progress_bar.pack(pady=(0, 10))
        
        # Interpretation
        self.cosine_interpretation_label = ttk.Label(self.cosine_frame, text="", 
                                                    font=("Arial", 10), wraplength=500)
        self.cosine_interpretation_label.pack(pady=(0, 10))
        
        # Explanation
        explanation = ("Cosine Similarity measures the angle between character frequency vectors.\n"
                      "Formula: (A · B) / (||A|| × ||B||) where A and B are character frequency vectors.")
        ttk.Label(self.cosine_frame, text=explanation, font=("Arial", 9), 
                wraplength=500, foreground="gray").pack()
    
    def create_edit_tab(self):
        """Create the Edit Distance tab"""
        # Score display
        score_frame = ttk.Frame(self.edit_frame)
        score_frame.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(score_frame, text="Edit Distance Similarity:").pack(side=tk.LEFT)
        self.edit_score_label = ttk.Label(score_frame, text="--", font=("Arial", 12, "bold"))
        self.edit_score_label.pack(side=tk.LEFT, padx=(10, 0))
        
        # Progress bar
        self.edit_progress_var = tk.DoubleVar()
        self.edit_progress_bar = ttk.Progressbar(self.edit_frame, variable=self.edit_progress_var, 
                                                maximum=1.0, length=300)
        self.edit_progress_bar.pack(pady=(0, 10))
        
        # Interpretation
        self.edit_interpretation_label = ttk.Label(self.edit_frame, text="", 
                                                  font=("Arial", 10), wraplength=500)
        self.edit_interpretation_label.pack(pady=(0, 10))
        
        # Explanation
        explanation = ("Edit Distance Similarity measures how many operations (insert, delete, substitute)\n"
                      "are needed to transform one word into another. Higher score = more similar.")
        ttk.Label(self.edit_frame, text=explanation, font=("Arial", 9), 
                wraplength=500, foreground="gray").pack()
    
    def calculate_similarity(self):
        """Calculate similarity using multiple methods"""
        word1 = self.word1_entry.get().strip().lower()
        word2 = self.word2_entry.get().strip().lower()
        
        if not word1 or not word2:
            messagebox.showwarning("Warning", "Please enter both words")
            return
        
        try:
            # Calculate Jaccard similarity
            jaccard_score = self.jaccard_similarity(word1, word2)
            self.jaccard_score_label.config(text=f"{jaccard_score:.4f}")
            self.jaccard_progress_var.set(jaccard_score)
            self.jaccard_interpretation_label.config(text=self.get_interpretation(jaccard_score))
            
            # Calculate Cosine similarity
            cosine_score = self.cosine_similarity(word1, word2)
            self.cosine_score_label.config(text=f"{cosine_score:.4f}")
            self.cosine_progress_var.set(cosine_score)
            self.cosine_interpretation_label.config(text=self.get_interpretation(cosine_score))
            
            # Calculate Edit Distance similarity
            edit_score = self.edit_distance_similarity(word1, word2)
            self.edit_score_label.config(text=f"{edit_score:.4f}")
            self.edit_progress_var.set(edit_score)
            self.edit_interpretation_label.config(text=self.get_interpretation(edit_score))
            
        except Exception as e:
            messagebox.showerror("Error", f"Error calculating similarity: {str(e)}")
    
    def jaccard_similarity(self, word1: str, word2: str) -> float:
        """Calculate Jaccard similarity between two words"""
        set1 = set(word1)
        set2 = set(word2)
        
        intersection = len(set1.intersection(set2))
        union = len(set1.union(set2))
        
        return intersection / union if union > 0 else 0.0
    
    def cosine_similarity(self, word1: str, word2: str) -> float:
        """Calculate Cosine similarity between two words using character frequencies"""
        # Create character frequency vectors
        freq1 = Counter(word1)
        freq2 = Counter(word2)
        
        # Get all unique characters
        all_chars = set(freq1.keys()).union(set(freq2.keys()))
        
        # Create vectors
        vec1 = [freq1.get(char, 0) for char in all_chars]
        vec2 = [freq2.get(char, 0) for char in all_chars]
        
        # Calculate cosine similarity
        dot_product = sum(a * b for a, b in zip(vec1, vec2))
        norm1 = sum(a * a for a in vec1) ** 0.5
        norm2 = sum(b * b for b in vec2) ** 0.5
        
        if norm1 == 0 or norm2 == 0:
            return 0.0
        
        return dot_product / (norm1 * norm2)
    
    def edit_distance_similarity(self, word1: str, word2: str) -> float:
        """Calculate similarity based on edit distance"""
        def levenshtein_distance(s1, s2):
            if len(s1) < len(s2):
                return levenshtein_distance(s2, s1)
            
            if len(s2) == 0:
                return len(s1)
            
            previous_row = list(range(len(s2) + 1))
            for i, c1 in enumerate(s1):
                current_row = [i + 1]
                for j, c2 in enumerate(s2):
                    insertions = previous_row[j + 1] + 1
                    deletions = current_row[j] + 1
                    substitutions = previous_row[j] + (c1 != c2)
                    current_row.append(min(insertions, deletions, substitutions))
                previous_row = current_row
            
            return previous_row[-1]
        
        distance = levenshtein_distance(word1, word2)
        max_len = max(len(word1), len(word2))
        
        # Convert distance to similarity (0 = identical, 1 = completely different)
        similarity = 1 - (distance / max_len) if max_len > 0 else 1.0
        return similarity
    
    def get_interpretation(self, score: float) -> str:
        """Get human-readable interpretation of the similarity score"""
        if score >= 0.8:
            return "Very High Similarity: These words are very closely related."
        elif score >= 0.6:
            return "High Similarity: These words are quite similar."
        elif score >= 0.4:
            return "Moderate Similarity: These words have some similarity."
        elif score >= 0.2:
            return "Low Similarity: These words have minimal similarity."
        else:
            return "Very Low Similarity: These words are quite different."
    
    def load_example(self, word1: str, word2: str):
        """Load example words into the input fields"""
        self.word1_entry.delete(0, tk.END)
        self.word1_entry.insert(0, word1)
        self.word2_entry.delete(0, tk.END)
        self.word2_entry.insert(0, word2)
        
        # Auto-calculate
        self.calculate_similarity()

def main():
    root = tk.Tk()
    app = SimpleSemanticSimilarityGUI(root)
    
    # Center the window
    root.update_idletasks()
    x = (root.winfo_screenwidth() // 2) - (root.winfo_width() // 2)
    y = (root.winfo_screenheight() // 2) - (root.winfo_height() // 2)
    root.geometry(f"+{x}+{y}")
    
    root.mainloop()

if __name__ == "__main__":
    main() 