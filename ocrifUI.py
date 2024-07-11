import os
import pytesseract
from PIL import Image
from pdf2image import convert_from_path
from docx import Document
from tkinter import Tk, Label, Button, filedialog

# Function to check if a corresponding docx file exists
def docx_exists(pdf_path):
    docx_path = os.path.splitext(pdf_path)[0] + '.docx'
    return os.path.exists(docx_path)

# Function to convert PDF to Word document only if a corresponding docx file doesn't exist
def convert_pdf_to_word(pdf_path):
    if not docx_exists(pdf_path):
        # Convert the PDF to images
        images = convert_from_path(pdf_path)

        # Create a new Word document
        doc = Document()

        # Loop through each image and extract text using OCR
        for i, image in enumerate(images):
            # Save the image as a temporary file
            image_path = f'temp_image_{i}.jpg'
            image.save(image_path, 'JPEG')

            # Extract text using OCR
            text = pytesseract.image_to_string(Image.open(image_path), lang='fas+equ')

            # Add the extracted text to the Word document
            doc.add_paragraph(text)

        # Save the Word document with the same name as the PDF file
        docx_path = os.path.splitext(pdf_path)[0] + '.docx'
        doc.save(docx_path)

# Function to handle file selection and conversion
def select_and_convert():
    file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    if file_path:
        if not docx_exists(file_path):
            convert_pdf_to_word(file_path)
            result_label.config(text=f"Conversion complete for {file_path}.")
        else:
            result_label.config(text=f"Docx file already exists for {file_path}. Skipping OCR.")

# Create UI
root = Tk()
root.title("PDF to Docx Converter")

# UI elements
label = Label(root, text="Select a PDF file to convert:")
label.pack(pady=10)

convert_button = Button(root, text="Select and Convert", command=select_and_convert)
convert_button.pack(pady=10)

result_label = Label(root, text="")
result_label.pack(pady=10)

# Run the UI loop
root.mainloop()
