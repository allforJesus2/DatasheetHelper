import os
from pathlib import Path
from typing import List, Dict, Generator
import pypdf
import docx
import pandas as pd
from sentence_transformers import SentenceTransformer
import chromadb
import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext
import threading
from datetime import datetime
import nltk
from nltk.tokenize import sent_tokenize
import json

nltk.download('punkt', quiet=True)


class JsonResultsViewer:
    def __init__(self, results, parent=None):
        self.window = tk.Toplevel(parent)
        self.window.title("Search Results Viewer")
        self.window.geometry("800x600")
        self.results = results
        self.setup_gui()

    def setup_gui(self):
        # Main frame with tree and text widgets
        paned = ttk.PanedWindow(self.window, orient='horizontal')
        paned.pack(fill='both', expand=True, padx=5, pady=5)

        # Left side - Treeview for files
        tree_frame = ttk.Frame(paned)
        tree_frame.pack(fill='both', expand=True)

        self.tree = ttk.Treeview(tree_frame, selectmode='browse')
        self.tree.heading('#0', text='Files')
        self.tree.pack(fill='both', expand=True)
        self.tree.bind('<<TreeviewSelect>>', self.on_select)

        # Right side - Content display
        content_frame = ttk.Frame(paned)
        content_frame.pack(fill='both', expand=True)

        # Metadata section
        self.metadata_text = scrolledtext.ScrolledText(content_frame, height=5)
        self.metadata_text.pack(fill='x', padx=5, pady=5)

        # Content sections
        self.content_text = scrolledtext.ScrolledText(content_frame, height=20)
        self.content_text.pack(fill='both', expand=True, padx=5, pady=5)

        # Open file button
        ttk.Button(content_frame, text="Open File", command=self.open_file).pack(pady=5)

        paned.add(tree_frame)
        paned.add(content_frame)

        self.populate_tree()

    def populate_tree(self):
        for filename in self.results.keys():
            self.tree.insert('', 'end', text=filename, iid=filename)

    def on_select(self, event):
        selected = self.tree.selection()
        if not selected:
            return

        filename = selected[0]
        file_data = self.results[filename]

        # Display metadata
        self.metadata_text.delete(1.0, tk.END)
        metadata = file_data['metadata']
        metadata_str = "\n".join(f"{k}: {v}" for k, v in metadata.items())
        self.metadata_text.insert('1.0', f"Metadata:\n{metadata_str}")

        # Display chunks
        self.content_text.delete(1.0, tk.END)
        for i, chunk in enumerate(file_data['chunks'], 1):
            self.content_text.insert('end', f"\nSection {i}:\n{chunk}\n")
            self.content_text.insert('end', '-' * 50 + '\n')

    def open_file(self):
        selected = self.tree.selection()
        if not selected:
            return

        filename = selected[0]
        file_path = self.results[filename]['metadata']['path']
        try:
            os.startfile(file_path)  # Windows
        except AttributeError:
            try:
                import subprocess
                subprocess.call(('xdg-open', file_path))  # Linux
            except:
                try:
                    subprocess.call(('open', file_path))  # MacOS
                except:
                    tk.messagebox.showerror("Error", "Could not open file")

class DocumentChunker:
    def __init__(self, chunk_size: int = 100, chunk_overlap: int = 10):
        self.chunk_size = chunk_size
        self.chunk_overlap = chunk_overlap

    def create_chunks(self, text: str) -> Generator[str, None, None]:
        sentences = sent_tokenize(text)
        current_chunk = []
        current_length = 0

        for sentence in sentences:
            sentence_length = len(sentence)

            if current_length + sentence_length > self.chunk_size:
                if current_chunk:
                    yield " ".join(current_chunk)
                    overlap_size = 0
                    overlap_chunk = []
                    for s in reversed(current_chunk):
                        if overlap_size + len(s) < self.chunk_overlap:
                            overlap_chunk.insert(0, s)
                            overlap_size += len(s)
                        else:
                            break
                    current_chunk = overlap_chunk
                    current_length = overlap_size

            current_chunk.append(sentence)
            current_length += sentence_length

        if current_chunk:
            yield " ".join(current_chunk)

class DocumentVectorizerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Document Vectorizer")
        self.root.geometry("800x600")
        self.output_format = tk.StringVar(value="text")
        self.n_results = tk.IntVar(value=5)
        self.sections_limit = tk.IntVar(value=3)
        # Add search option variables
        self.search_content = tk.BooleanVar(value=True)
        self.search_filename = tk.BooleanVar(value=False)
        self.search_metadata = tk.BooleanVar(value=False)
        self.setup_gui()
        self.vectorizer = None


    def setup_gui(self):
        # Database selection
        db_frame = ttk.Frame(self.root)
        db_frame.pack(fill='x', padx=5, pady=5)

        self.db_path = tk.StringVar(value="./chroma_db")
        ttk.Label(db_frame, text="Database:").pack(side='left')
        ttk.Entry(db_frame, textvariable=self.db_path).pack(side='left', fill='x', expand=True, padx=5)
        ttk.Button(db_frame, text="Browse DB", command=self.browse_db).pack(side='left')
        ttk.Button(db_frame, text="Load DB", command=self.load_database).pack(side='left', padx=5)

        # Chunking options
        chunk_frame = ttk.LabelFrame(self.root, text="Chunking Options")
        chunk_frame.pack(fill='x', padx=5, pady=5)

        self.chunk_size = tk.IntVar(value=500)
        ttk.Label(chunk_frame, text="Chunk Size:").pack(side='left')
        ttk.Entry(chunk_frame, textvariable=self.chunk_size, width=10).pack(side='left', padx=5)

        self.chunk_overlap = tk.IntVar(value=100)
        ttk.Label(chunk_frame, text="Overlap:").pack(side='left')
        ttk.Entry(chunk_frame, textvariable=self.chunk_overlap, width=10).pack(side='left', padx=5)

        # Folder selection and collection
        folder_frame = ttk.Frame(self.root)
        folder_frame.pack(fill='x', padx=5, pady=5)

        self.folder_path = tk.StringVar()
        ttk.Label(folder_frame, text="Documents:").pack(side='left')
        ttk.Entry(folder_frame, textvariable=self.folder_path).pack(side='left', fill='x', expand=True, padx=5)
        ttk.Button(folder_frame, text="Browse", command=self.browse_folder).pack(side='left')

        self.collection_name = tk.StringVar(value="documents")
        ttk.Label(folder_frame, text="Collection:").pack(side='left', padx=5)
        self.collections_dropdown = ttk.Combobox(folder_frame, textvariable=self.collection_name, state='readonly')
        self.collections_dropdown.pack(side='left')
        self.collections_dropdown.bind('<<ComboboxSelected>>', self.on_collection_select)

        # Process button and progress
        self.process_btn = ttk.Button(self.root, text="Process Documents", command=self.start_processing)
        self.process_btn.pack(pady=5)

        self.progress_var = tk.StringVar(value="Ready")
        ttk.Label(self.root, textvariable=self.progress_var).pack(pady=5)
        # Output format selection
        format_frame = ttk.Frame(self.root)
        format_frame.pack(fill='x', padx=5, pady=5)

        ttk.Label(format_frame, text="Output Format:").pack(side='left')
        ttk.Radiobutton(format_frame, text="Text", variable=self.output_format, value="text").pack(side='left', padx=5)
        ttk.Radiobutton(format_frame, text="JSON", variable=self.output_format, value="json").pack(side='left')

        # Add search options frame
        options_frame = ttk.LabelFrame(self.root, text="Search Options")
        options_frame.pack(fill='x', padx=5, pady=5)

        ttk.Checkbutton(options_frame, text="Search Content", variable=self.search_content).pack(side='left', padx=5)
        ttk.Checkbutton(options_frame, text="Search Filename", variable=self.search_filename).pack(side='left', padx=5)
        ttk.Checkbutton(options_frame, text="Search Metadata", variable=self.search_metadata).pack(side='left', padx=5)
        self.exact_match = tk.BooleanVar(value=True)
        ttk.Checkbutton(options_frame, text="Exact Match", variable=self.exact_match).pack(side='left', padx=5)

        # Add limits frame
        limits_frame = ttk.LabelFrame(self.root, text="Search Limits")
        limits_frame.pack(fill='x', padx=5, pady=5)

        ttk.Label(limits_frame, text="Max Results:").pack(side='left')
        ttk.Entry(limits_frame, textvariable=self.n_results, width=5).pack(side='left', padx=5)

        ttk.Label(limits_frame, text="Sections per Result:").pack(side='left')
        ttk.Entry(limits_frame, textvariable=self.sections_limit, width=5).pack(side='left', padx=5)

        # Search frame
        search_frame = ttk.Frame(self.root)
        search_frame.pack(fill='x', padx=5, pady=5)

        self.search_query = tk.StringVar()
        search_entry = ttk.Entry(search_frame, textvariable=self.search_query)
        search_entry.pack(side='left', fill='x', expand=True)
        search_entry.bind('<Return>', lambda event: self.perform_search())  # Bind Enter key
        ttk.Button(search_frame, text="Search", command=self.perform_search).pack(side='left', padx=5)
        ttk.Button(search_frame, text="View Results", command=self.view_results).pack(side='left', padx=5)
        ttk.Button(search_frame, text="Save Results", command=self.save_results).pack(side='left')

        # Results
        self.results_text = scrolledtext.ScrolledText(self.root, height=20)
        self.results_text.pack(fill='both', expand=True, padx=5, pady=5)

        # Store the last search results
        self.last_results = None

    def should_process_file(self, file_path: Path) -> bool:
        # Skip common irrelevant files
        if file_path.name.startswith('~$') or file_path.name.startswith('.'):
            return False

        # Skip files over size limit (e.g. 10MB)
        if file_path.stat().st_size > 10_000_000:
            return False

        # Skip backup/temp files
        skip_patterns = ['backup', 'old', 'archive', 'temp']
        if any(pattern in file_path.name.lower() for pattern in skip_patterns):
            return False

        return True


    def view_results(self):
        if not self.last_results:
            self.log_message("No results to view")
            return
        JsonResultsViewer(self.last_results, self.root)

    def browse_db(self):
        db_path = filedialog.askdirectory(title="Select Database Directory")
        if db_path:
            self.db_path.set(db_path)

    def load_database(self):
        try:
            db_path = self.db_path.get()
            self.vectorizer = DocumentVectorizer(
                folder_path="",
                db_path=db_path,
                collection_name=self.collection_name.get(),
                chunk_size=self.chunk_size.get(),
                chunk_overlap=self.chunk_overlap.get()
            )
            collections = self.vectorizer.chroma_client.list_collections()
            collection_names = [c.name for c in collections]
            self.collections_dropdown['values'] = collection_names
            if collection_names:
                self.collections_dropdown.set(collection_names[0])
                self.collection_name.set(collection_names[0])
            self.log_message(f"Loaded database from {db_path}")
        except Exception as e:
            self.log_message(f"Error loading database: {str(e)}")

    def on_collection_select(self, event):
        selected = self.collections_dropdown.get()
        if self.vectorizer:
            self.vectorizer.set_collection(selected)
            self.log_message(f"Switched to collection: {selected}")

    def browse_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.folder_path.set(folder)

    def log_message(self, message):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.results_text.insert('end', f"[{timestamp}] {message}\n")
        self.results_text.see('end')

    def start_processing(self):
        if not self.folder_path.get():
            self.log_message("Please select a folder first")
            return

        self.process_btn.config(state='disabled')
        self.progress_var.set("Processing...")

        def process():
            try:
                if not self.vectorizer:
                    self.vectorizer = DocumentVectorizer(
                        self.folder_path.get(),
                        db_path=self.db_path.get(),
                        collection_name=self.collection_name.get(),
                        chunk_size=self.chunk_size.get(),
                        chunk_overlap=self.chunk_overlap.get()
                    )
                else:
                    self.vectorizer.folder_path = Path(self.folder_path.get())
                self.vectorizer.process_documents(callback=self.log_message)
                # Remove the database reload since it's no longer needed
                self.root.after(0, lambda: self.progress_var.set("Ready"))
                self.root.after(0, lambda: self.process_btn.config(state='normal'))
            except Exception as e:
                self.root.after(0, lambda: self.log_message(f"Error: {str(e)}"))
                self.root.after(0, lambda: self.progress_var.set("Error"))
                self.root.after(0, lambda: self.process_btn.config(state='normal'))

        threading.Thread(target=process, daemon=True).start()
    def save_results(self):
        if not self.last_results:
            self.log_message("No results to save")
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("Text files", "*.txt"), ("All files", "*.*")]
        )
        if file_path:
            try:
                if self.output_format.get() == "json":
                    with open(file_path, 'w', encoding='utf-8') as f:
                        json.dump(self.last_results, f, indent=2, ensure_ascii=False)
                else:
                    with open(file_path, 'w', encoding='utf-8') as f:
                        f.write(self.results_text.get(1.0, tk.END))
                self.log_message(f"Results saved to {file_path}")
            except Exception as e:
                self.log_message(f"Error saving results: {str(e)}")

    def display_results(self, results: Dict, format_type: str):
        self.last_results = results
        self.results_text.delete(1.0, tk.END)

        if format_type == "json":
            json_str = json.dumps(results, indent=2, ensure_ascii=False)
            self.results_text.insert('end', json_str)
        else:
            for filename, data in results.items():
                metadata = data['metadata']
                chunks = data['chunks']

                self.results_text.insert('end', f"\nDocument: {filename}\n")
                self.results_text.insert('end', f"Path: {metadata['path']}\n")
                self.results_text.insert('end', "Relevant sections:\n")

                for i, chunk in enumerate(chunks, 1):
                    self.results_text.insert('end', f"\nSection {i}:\n{chunk}\n")

                self.results_text.insert('end', "-" * 50 + "\n")

    def perform_search(self):
        if not self.vectorizer:
            self.log_message("Please load or create a database first")
            return

        query = self.search_query.get()
        if not query:
            self.log_message("Please enter a search query")
            return

        try:
            results = self.vectorizer.advanced_search(
                query,
                n_results=self.n_results.get(),
                search_content=self.search_content.get(),
                search_filename=self.search_filename.get(),
                search_metadata=self.search_metadata.get(),
                exact_match=self.exact_match.get()
            )
            # Limit sections per result
            for doc in results.values():
                doc['chunks'] = doc['chunks'][:self.sections_limit.get()]
            self.display_results(results, self.output_format.get())
        except Exception as e:
            self.log_message(f"Search error: {str(e)}")

class DocumentVectorizer:
    def __init__(self, folder_path: str, db_path: str = "./chroma_db",
                 collection_name: str = "documents", chunk_size: int = 250,  # Reduced default chunk size
                 chunk_overlap: int = 50):  # Reduced default overlap
        self.folder_path = Path(folder_path) if folder_path else None
        self.model = SentenceTransformer('sentence-transformers/all-MiniLM-L6-v2')
        self.chroma_client = chromadb.PersistentClient(path=db_path)
        self.collection = self.chroma_client.get_or_create_collection(
            name=collection_name,
            metadata={"hnsw:space": "cosine"}  # Explicitly set distance metric
        )
        self.chunker = DocumentChunker(chunk_size, chunk_overlap)

    def set_collection(self, collection_name: str):
        self.collection = self.chroma_client.get_or_create_collection(collection_name)

    def extract_text_from_pdf(self, file_path: Path) -> str:
        text = ""
        with open(file_path, 'rb') as file:
            pdf = pypdf.PdfReader(file)
            for page in pdf.pages:
                text += page.extract_text() + "\n"
        return text

    def extract_text_from_docx(self, file_path: Path) -> str:
        doc = docx.Document(file_path)
        return "\n".join([paragraph.text for paragraph in doc.paragraphs])

    def extract_text_from_xlsx(self, file_path: Path) -> str:
        df = pd.read_excel(file_path)
        return df.to_string()

    def extract_text(self, file_path: Path) -> str:
        if file_path.suffix.lower() == '.pdf':
            return self.extract_text_from_pdf(file_path)
        elif file_path.suffix.lower() == '.docx':
            return self.extract_text_from_docx(file_path)
        elif file_path.suffix.lower() == '.xlsx' or file_path.suffix.lower() == '.xlsm':
            return self.extract_text_from_xlsx(file_path)
        return ""

    def process_documents(self, callback=None) -> None:
        if not self.folder_path:
            raise ValueError("No folder path specified")

        # Get existing document metadata to check modification times
        existing_docs = {}
        if self.collection.count() > 0:
            collection_data = self.collection.get()
            for doc_id, metadata in zip(collection_data['ids'], collection_data['metadatas']):
                file_path = metadata['path']
                chunk_index = metadata['chunk_index']
                if chunk_index == 0:  # Only store first chunk's metadata for each file
                    existing_docs[file_path] = {
                        'modified_time': metadata.get('modified_time'),
                        'doc_ids': [id for id in collection_data['ids']
                                    if id.startswith(file_path + '#chunk')]
                    }

        for file_path in self.folder_path.glob('**/*'):
            if file_path.is_file():
                try:
                    str_path = str(file_path)
                    current_mtime = datetime.fromtimestamp(os.path.getmtime(file_path)).strftime('%Y-%m-%d %H:%M:%S')

                    # Check if file exists in database and if it's been modified
                    needs_processing = True
                    if str_path in existing_docs:
                        old_mtime = existing_docs[str_path]['modified_time']
                        if old_mtime == current_mtime:
                            needs_processing = False
                            if callback:
                                callback(f"Skipped: {file_path} (not modified)")

                    if needs_processing:
                        if callback:
                            callback(f"Processing: {file_path}")

                        # Delete existing chunks for this file if they exist
                        if str_path in existing_docs:
                            self.collection.delete(ids=existing_docs[str_path]['doc_ids'])
                            if callback:
                                callback(f"Deleted existing chunks for: {file_path}")

                        # Process and add new chunks
                        text = self.extract_text(file_path)
                        if text:
                            chunks = list(self.chunker.create_chunks(text))
                            chunk_ids = [f"{file_path}#chunk{i}" for i in range(len(chunks))]
                            metadatas = [{
                                "filename": file_path.name,
                                "path": str_path,
                                "file_type": file_path.suffix.lower(),
                                "size": os.path.getsize(file_path),
                                "chunk_index": i,
                                "total_chunks": len(chunks),
                                "modified_time": current_mtime,
                                "processed_time": datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                            } for i in range(len(chunks))]

                            self.collection.add(
                                documents=chunks,
                                metadatas=metadatas,
                                ids=chunk_ids
                            )
                            if callback:
                                callback(f"Processed: {file_path} ({len(chunks)} chunks)")

                except Exception as e:
                    if callback:
                        callback(f"Error processing {file_path}: {str(e)}")

    def advanced_search(self, query: str, n_results: int = 5,
                        search_content: bool = True,
                        search_filename: bool = False,
                        search_metadata: bool = False,
                        exact_match: bool = True) -> Dict:
        """
        Advanced search method that supports exact matching, content search, and metadata search.

        Args:
            query (str): The search query
            n_results (int): Maximum number of results to return
            search_content (bool): Whether to search in document content
            search_filename (bool): Whether to search in filenames
            search_metadata (bool): Whether to search in metadata
            exact_match (bool): Whether to perform exact string matching

        Returns:
            Dict: Dictionary of search results organized by filename
        """
        all_results = {}

        # Get all documents if we need them for exact matching or metadata search
        if exact_match or search_filename or search_metadata:
            all_docs = self.collection.get()

        if search_content:
            if exact_match:
                # Perform exact string matching
                if all_docs and all_docs['documents']:
                    for doc, metadata, doc_id in zip(all_docs['documents'],
                                                     all_docs['metadatas'],
                                                     all_docs['ids']):
                        # Case-insensitive exact match
                        if query.lower() in doc.lower():
                            filename = metadata['filename']
                            if filename not in all_results:
                                all_results[filename] = {
                                    'chunks': [doc],
                                    'metadata': metadata
                                }
                            elif doc not in all_results[filename]['chunks']:
                                all_results[filename]['chunks'].append(doc)

            # If we don't have enough exact matches or exact match is disabled
            if len(all_results) < n_results and not exact_match:
                semantic_results = self.collection.query(
                    query_texts=[query],
                    n_results=n_results
                )
                # Process semantic search results
                if semantic_results and semantic_results['documents']:
                    for doc, metadata in zip(semantic_results['documents'][0],
                                             semantic_results['metadatas'][0]):
                        filename = metadata['filename']
                        if filename not in all_results:
                            all_results[filename] = {
                                'chunks': [doc],
                                'metadata': metadata
                            }
                        elif doc not in all_results[filename]['chunks']:
                            all_results[filename]['chunks'].append(doc)

        # Handle filename and metadata searching
        if (search_filename or search_metadata) and all_docs:
            for doc, metadata, doc_id in zip(all_docs['documents'],
                                             all_docs['metadatas'],
                                             all_docs['ids']):
                filename = metadata['filename']
                should_add = False

                # Search in filename
                if search_filename and query.lower() in filename.lower():
                    should_add = True

                # Search in metadata
                if search_metadata:
                    metadata_to_search = {
                        **metadata,
                        'file_size_formatted': f"{metadata['size'] / 1024:.2f} KB",
                    }
                    # Check all metadata values for the query
                    if any(query.lower() in str(v).lower() for v in metadata_to_search.values()):
                        should_add = True

                if should_add:
                    if filename not in all_results:
                        all_results[filename] = {
                            'chunks': [doc],
                            'metadata': metadata
                        }
                    elif doc not in all_results[filename]['chunks']:
                        all_results[filename]['chunks'].append(doc)

        # Sort results by number of matching chunks (most relevant first)
        sorted_results = dict(sorted(
            all_results.items(),
            key=lambda x: len(x[1]['chunks']),
            reverse=True
        )[:n_results])

        return sorted_results
    def _process_results(self, results, all_results):
        for doc, metadata in zip(results['documents'][0], results['metadatas'][0]):
            filename = metadata['filename']
            if filename not in all_results:
                all_results[filename] = {
                    'chunks': [],
                    'metadata': metadata
                }
            all_results[filename]['chunks'].append(doc)



def main():
    root = tk.Tk()
    app = DocumentVectorizerGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()