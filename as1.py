import os
import glob
from pdfminer.high_level import extract_text
def extract_text_from_pdf_directory(directory_path):
    # Get all PDF file paths in the directory
    pdf_files = glob.glob(os.path.join(directory_path, "*.pdf"))

    # Iterate over each PDF file
    for pdf_file in pdf_files:
        # Extract text from the PDF file
        try:
            text = extract_text(pdf_file)
        except Exception as e:
            print(f"Error extracting text from {pdf_file}: {e}")
            continue

        # Create a text file path by replacing the extension with '.txt'
        txt_file = os.path.splitext(pdf_file)[0] + ".txt"

        # Save the extracted text into a separate text file
        with open(txt_file, "w", encoding="utf-8") as f:
            f.write(text)

        print(f"Text extracted from {pdf_file} and saved to {txt_file}.")

# Usage example
directory_path = "C:\\dir"
extract_text_from_pdf_directory(directory_path)
