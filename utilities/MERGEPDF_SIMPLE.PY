#!/usr/bin/env python3
"""
Simple PDF Merger Script
Merges multiple PDF files into a single output file.
"""

import os
from PyPDF2 import PdfMerger

def merge_pdfs(input_files, output_file):
    """
    Merge multiple PDF files into a single output file.
    
    Args:
        input_files (list): List of input PDF file paths
        output_file (str): Output PDF file path
    """
    if len(input_files) < 2:
        print("Error: At least 2 PDF files are required for merging.")
        return False
    
    # Check if all input files exist
    for file_path in input_files:
        if not os.path.exists(file_path):
            print(f"Error: File does not exist: {file_path}")
            return False
    
    try:
        merger = PdfMerger()
        
        print(f"Starting to merge {len(input_files)} PDF files...")
        
        for pdf_file in input_files:
            print(f"  Adding: {os.path.basename(pdf_file)}")
            merger.append(pdf_file)
        
        # Create output directory if it doesn't exist
        output_dir = os.path.dirname(output_file)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir)
        
        # Write the merged PDF
        print(f"Writing merged PDF to: {output_file}")
        with open(output_file, 'wb') as output_pdf:
            merger.write(output_pdf)
        
        merger.close()
        
        print(f"\nSuccess! Merged {len(input_files)} PDF files into '{output_file}'")
        return True
        
    except Exception as e:
        print(f"Error merging PDFs: {e}")
        return False

# Example usage
if __name__ == "__main__":
    # List of PDF files to merge (change these to your actual files)
    pdf_files = [
        r"S:\Projects\24317_Electra_CO_EPCM\Engineering\Instrumentation\DataSheets\VALVES\PDFS\SWISS PLUG COVER.pdf",
        r"S:\Projects\24317_Electra_CO_EPCM\Engineering\Instrumentation\DataSheets\VALVES\PDFS\SWISS PLUG DS.pdf"
    ]
    
    # Output file name
    output_file = r"S:\Projects\24317_Electra_CO_EPCM\Engineering\Instrumentation\DataSheets\VALVES\PDFS\SWISS PLUG combined.pdf"
    
    # Merge the PDFs
    merge_pdfs(pdf_files, output_file)
