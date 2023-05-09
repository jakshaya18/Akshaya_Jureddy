# Import the necessary libraries
import os
import io
import PyPDF2
from docx2pdf import convert

# Specify the file path for the PDF file to be converted
pdf_path = "path/to/pdf/file.pdf"

# Read the PDF file
with open(pdf_path, "rb") as pdf_file:
    pdf_reader = PyPDF2.PdfFileReader(pdf_file)
    num_pages = pdf_reader.numPages
    text = ""
    
    # Loop through each page of the PDF and extract the text
    for page in range(num_pages):
        pdf_page = pdf_reader.getPage(page)
        text += pdf_page.extractText()
    
    # Specify the file path for the Word document to be created
    word_path = "path/to/word/document.docx"
    
    # Write the text to a Word document
    with io.open(word_path, "w", encoding="utf-8") as word_file:
        word_file.write(text)
    
    # Convert the Word document to PDF format
    convert(word_path)

# Delete the temporary Word document
os.remove(word_path)

