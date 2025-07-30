import os
# Set environment variable to handle OpenMP runtime conflict
os.environ['KMP_DUPLICATE_LIB_OK'] = 'TRUE'

import tkinter as tk
from tkinter import ttk, messagebox
from sentence_transformers import SentenceTransformer
import numpy as np
from typing import Tuple
import threading

class SemanticSimilarityGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Semantic Similarity Calculator")
        self.root.geometry("600x400")
        self.root.resizable(True, True)
        
        # Initialize the model (will be loaded in background)
        self.model = None
        self.model_loaded = False
        self.loading_model = False
        
        # Create the main frame
        self.main_frame = ttk.Frame(root, padding="20")
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Create widgets
        self.create_widgets()
        
        # Load model in background
        self.load_model_async()
    
    def create_widgets(self):
        # Title
        title_label = ttk.Label(self.main_frame, text="Semantic Similarity Calculator", 
                               font=("Arial", 16, "bold"))
        title_label.pack(pady=(0, 20))
        
        # Model status
        self.status_label = ttk.Label(self.main_frame, text="Loading model...", 
                                     foreground="orange")
        self.status_label.pack(pady=(0, 20))
        
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
                                          command=self.calculate_similarity, state="disabled")
        self.calculate_button.pack(pady=(10, 0))
        
        # Results frame
        results_frame = ttk.LabelFrame(self.main_frame, text="Results", padding="10")
        results_frame.pack(fill=tk.BOTH, expand=True)
        
        # Similarity score
        score_frame = ttk.Frame(results_frame)
        score_frame.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(score_frame, text="Similarity Score:").pack(side=tk.LEFT)
        self.score_label = ttk.Label(score_frame, text="--", font=("Arial", 12, "bold"))
        self.score_label.pack(side=tk.LEFT, padx=(10, 0))
        
        # Progress bar for score visualization
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(results_frame, variable=self.progress_var, 
                                           maximum=1.0, length=300)
        self.progress_bar.pack(pady=(0, 10))
        
        # Interpretation
        self.interpretation_label = ttk.Label(results_frame, text="", 
                                             font=("Arial", 10), wraplength=500)
        self.interpretation_label.pack(pady=(0, 10))
        
        # Examples frame
        examples_frame = ttk.LabelFrame(self.main_frame, text="Quick Examples", padding="10")
        examples_frame.pack(fill=tk.X, pady=(20, 0))
        
        # Example buttons
        examples = [
            ("cat", "dog"),
            ("happy", "sad"),
            ("computer", "laptop"),
            ("car", "bicycle"),
            ("apple", "orange")
        ]
        
        for i, (word1, word2) in enumerate(examples):
            btn = ttk.Button(examples_frame, text=f"{word1} vs {word2}", 
                           command=lambda w1=word1, w2=word2: self.load_example(w1, w2))
            btn.grid(row=i//3, column=i%3, padx=5, pady=2, sticky="ew")
        
        # Configure grid weights
        for i in range(3):
            examples_frame.columnconfigure(i, weight=1)
    
    def load_model_async(self):
        """Load the sentence transformer model in a background thread"""
        def load_model():
            try:
                self.loading_model = True
                # Use a lightweight model for faster loading
                self.model = SentenceTransformer('all-MiniLM-L6-v2')
                self.model_loaded = True
                self.loading_model = False
                
                # Update UI in main thread
                self.root.after(0, self.on_model_loaded)
            except Exception as e:
                self.loading_model = False
                self.root.after(0, lambda: self.on_model_error(str(e)))
        
        thread = threading.Thread(target=load_model, daemon=True)
        thread.start()
    
    def on_model_loaded(self):
        """Called when model is successfully loaded"""
        self.status_label.config(text="Model loaded successfully!", foreground="green")
        self.calculate_button.config(state="normal")
    
    def on_model_error(self, error_msg):
        """Called when model loading fails"""
        self.status_label.config(text=f"Error loading model: {error_msg}", foreground="red")
        messagebox.showerror("Error", f"Failed to load model: {error_msg}")
    
    def calculate_similarity(self):
        """Calculate semantic similarity between the two input words"""
        word1 = self.word1_entry.get().strip()
        word2 = self.word2_entry.get().strip()
        
        if not word1 or not word2:
            messagebox.showwarning("Warning", "Please enter both words")
            return
        
        if not self.model_loaded:
            messagebox.showwarning("Warning", "Model is still loading. Please wait.")
            return
        
        try:
            # Calculate similarity
            similarity_score = self.compute_similarity(word1, word2)
            
            # Update UI
            self.score_label.config(text=f"{similarity_score:.4f}")
            self.progress_var.set(similarity_score)
            
            # Update interpretation
            interpretation = self.get_interpretation(similarity_score)
            self.interpretation_label.config(text=interpretation)
            
        except Exception as e:
            messagebox.showerror("Error", f"Error calculating similarity: {str(e)}")
    
    def compute_similarity(self, word1: str, word2: str) -> float:
        """Compute cosine similarity between two words"""
        # Encode the words
        embeddings = self.model.encode([word1, word2])
        
        # Calculate cosine similarity
        similarity = np.dot(embeddings[0], embeddings[1]) / (
            np.linalg.norm(embeddings[0]) * np.linalg.norm(embeddings[1])
        )
        
        return float(similarity)
    
    def get_interpretation(self, score: float) -> str:
        """Get human-readable interpretation of the similarity score"""
        if score >= 0.8:
            return "Very High Similarity: These words are very closely related in meaning."
        elif score >= 0.6:
            return "High Similarity: These words are quite similar in meaning."
        elif score >= 0.4:
            return "Moderate Similarity: These words have some similarity in meaning."
        elif score >= 0.2:
            return "Low Similarity: These words have minimal similarity in meaning."
        else:
            return "Very Low Similarity: These words are quite different in meaning."
    
    def load_example(self, word1: str, word2: str):
        """Load example words into the input fields"""
        self.word1_entry.delete(0, tk.END)
        self.word1_entry.insert(0, word1)
        self.word2_entry.delete(0, tk.END)
        self.word2_entry.insert(0, word2)
        
        # Auto-calculate if model is loaded
        if self.model_loaded:
            self.calculate_similarity()

def main():
    root = tk.Tk()
    app = SemanticSimilarityGUI(root)
    
    # Center the window
    root.update_idletasks()
    x = (root.winfo_screenwidth() // 2) - (root.winfo_width() // 2)
    y = (root.winfo_screenheight() // 2) - (root.winfo_height() // 2)
    root.geometry(f"+{x}+{y}")
    
    root.mainloop()

if __name__ == "__main__":
    main() 