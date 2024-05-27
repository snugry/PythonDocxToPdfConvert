import os
from spire.doc import *
from spire.doc.common import *
import time

def convert_docx_to_pdf(docx_path, pdf_path):
    # Normalize the file paths
    #docx_path = os.path.normpath(docx_path)
    #pdf_path = os.path.normpath(pdf_path)

    # Add double quotes around file paths to handle spaces
    docx_path = f'{docx_path}'
    pdf_path = f'{pdf_path}'

    # Create a new Word Application object
    document = Document()
    time.sleep(1)

    try:
        # Open the source DOCX file
        document.LoadFromFile(docx_path)
        
        print("after open")
        time.sleep(2)
        # Save the file as PDF
        document.SaveToFile(pdf_path, FileFormat.PDF)
        document.Close()

        print(f"{docx_path} converted to PDF: {pdf_path}")
    except Exception as e:
        print(f"Error converting {docx_path} to PDF: {e}")

def convert_folder_to_pdf(input_folder, output_folder):
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    for file in os.listdir(input_folder):
        if file.endswith(".docx"):
            docx_path = os.path.join(input_folder, file)
            pdf_path = os.path.join(output_folder, f"{os.path.splitext(file)[0]}.pdf")
            print(pdf_path)
            convert_docx_to_pdf(docx_path, pdf_path)

if __name__ == "__main__":
    input_folder = "input\\folder"
    output_folder = "output\\folder"
    convert_folder_to_pdf(input_folder, output_folder)