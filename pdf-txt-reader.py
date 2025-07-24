import fitz

def read_pdf_as_text(pdf_file_path):
    # Open the PDF file using fitz
    doc = fitz.open(pdf_file_path)

    # Initialize an empty string to store the text
    text = ""

    # Iterate through each page in the PDF
    for page in doc:
        # Get the text from the page
        page_text = page.get_text()

        # Add the text to the overall text string
        text += page_text

    # Close the PDF file
    doc.close()

    # Return the extracted text
    return text

# Example usage:
pdf_file_path = r"C:\Users\dcaoili\Downloads\23321-01-P-I-013 Brice Barclay PO Rev 0.pdf"
text = read_pdf_as_text(pdf_file_path)
print(text)